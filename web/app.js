/* 秋招投递记录 · 前端 */
(() => {
  'use strict';

  const TEXT   = ['公司', '岗位', '薪资', '工作地点', '下一步动作', '备注'];
  const DATES  = ['投递日期', '截止/跟进时间', '笔试时间', '一面时间', '二面时间', 'HR面时间'];
  const SELECT = ['渠道', '状态', '面试形式', '结果'];
  const ALL    = [...TEXT, ...DATES, ...SELECT];
  const FINAL  = ['已发Offer', '已挂', '已放弃'];

  const $  = (id) => document.getElementById(id);
  const el = (n)  => $('f_' + n);
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g,
    c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  let records = [];
  let byRow = {};
  let options = {};
  let editing = false;
  let trash = [];
  let backups = [];

  /* ---------------- 工具 ---------------- */
  function toast(msg, kind = '') {
    const t = $('toast');
    t.textContent = msg;
    t.className = 'toast show ' + kind;
    clearTimeout(t._timer);
    t._timer = setTimeout(() => { t.className = 'toast ' + kind; }, kind === 'err' ? 6000 : 2600);
  }

  function todayStr() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  }

  function daysBetween(a, b) {   // a,b 为 'YYYY-MM-DD'
    return Math.round((new Date(a + 'T00:00:00') - new Date(b + 'T00:00:00')) / 86400000);
  }

  function statusClass(s) {
    if (s === '已发Offer') return 's3';
    if (s === '已挂' || s === '已放弃') return 's4';
    if (s && s.includes('面')) return 's2';
    if (s === '已投递') return 's0';
    return 's1';
  }

  async function api(path, body) {
    const opt = body
      ? { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) }
      : {};
    const res = await fetch(path, opt);
    let data;
    try { data = await res.json(); }
    catch { throw new Error('服务无响应，请确认启动窗口还开着'); }
    if (!res.ok || !data.ok) throw new Error(data.error || ('请求失败 ' + res.status));
    return data;
  }

  /* ---------------- 加载 ---------------- */
  async function load(keepTip) {
    try {
      const d = await api('/api/data');
      records = d.records;
      options = d.options;
      trash = d.trash || [];
      backups = d.backups || [];
      byRow = {};
      records.forEach(r => { byRow[r._row] = r; });
      fillOptions();
      renderStats(d.stats);
      renderTable();
      renderSafety();
      $('filePath').textContent = '数据文件：' + d.file;
      if (!keepTip) flash('✓ 已同步');
    } catch (e) {
      toast(e.message, 'err');
    }
  }

  function flash(text) {
    const s = $('saveTip');
    s.textContent = text;
    s.style.opacity = '1';
    setTimeout(() => { s.style.opacity = '0'; }, 1800);
  }

  /* ---------------- 回收站 / 历史版本 ---------------- */
  function fmtSize(n) {
    return n > 1048576 ? (n / 1048576).toFixed(1) + ' MB' : Math.round(n / 1024) + ' KB';
  }

  function renderSafety() {
    $('trashCount').textContent = '回收站 ' + trash.length;
    $('backupCount').textContent = '快照 ' + backups.length;

    $('trashEmpty').classList.toggle('hidden', trash.length > 0);
    $('trashBody').innerHTML = trash.map(t => `<tr>
      <td>${esc(t['删除时间'])}</td>
      <td class="co">${esc(t['公司'])}</td>
      <td class="pos">${esc(t['岗位'])}</td>
      <td>${esc(t['状态'])}</td>
      <td>${esc(t['投递日期'])}</td>
      <td><button class="op" data-act="untrash" data-trow="${t._trow}" type="button">还原</button></td>
    </tr>`).join('');

    $('backupEmpty').classList.toggle('hidden', backups.length > 0);
    $('backupList').innerHTML = backups.map(b => `<div class="backup-item">
      <span class="bt">${esc(b.time)}</span>
      <span class="bs">${fmtSize(b.size)}</span>
      <button class="op" data-act="restore" data-name="${esc(b.name)}" type="button">恢复到此版本</button>
    </div>`).join('');
  }

  function fillOptions() {
    SELECT.forEach(name => {
      const sel = el(name);
      const list = options[name] || [];
      const cur = sel.value;
      sel.innerHTML = '<option value=""></option>' +
        list.map(v => `<option value="${esc(v)}">${esc(v)}</option>`).join('');
      if (list.includes(cur)) sel.value = cur;
    });

    // 公司 datalist：历史公司 + 「选项」表里的公司库
    const cos = [...new Set([...(options._公司库 || []), ...records.map(r => r.公司).filter(Boolean)])];
    $('dl_公司').innerHTML = cos.map(v => `<option value="${esc(v)}">`).join('');
    const poss = [...new Set(records.map(r => r.岗位).filter(Boolean))];
    $('dl_岗位').innerHTML = poss.map(v => `<option value="${esc(v)}">`).join('');
  }

  /* ---------------- 统计卡片 ---------------- */
  function renderStats(s) {
    if (!s) return;
    const items = [
      ['总投递', s.total, ''],
      ['推进中', s.pending, 'acc'],
      ['进入面试', s.interview, ''],
      ['Offer', s.offer, 'ok'],
      ['已挂/放弃', s.dead, 'bad'],
      ['近7天投递', s.week, ''],
      ['逾期未跟进', s.overdue, s.overdue ? 'bad' : ''],
      ['Offer率', s.rate + '%', s.offer ? 'ok' : ''],
    ];
    $('cards').innerHTML = items.map(([k, v, c]) =>
      `<div class="card ${c}"><div class="k">${k}</div><div class="v">${v}</div></div>`).join('');
  }

  /* ---------------- 表格 ---------------- */
  function renderTable() {
    const q = $('search').value.trim().toLowerCase();
    const list = q ? records.filter(r =>
      ALL.some(n => (r[n] || '').toLowerCase().includes(q))) : records;

    $('count').textContent = list.length;
    $('empty').classList.toggle('hidden', list.length > 0);

    const today = todayStr();
    const opts = (name, cur) => '<option value=""></option>' +
      (options[name] || []).map(v =>
        `<option value="${esc(v)}"${v === cur ? ' selected' : ''}>${esc(v)}</option>`).join('');

    $('tbody').innerHTML = list.map(r => {
      const d = r['投递日期'];
      const days = d ? daysBetween(today, d) : '';
      const h = r['截止/跟进时间'];
      let nx = '';
      if (h) {
        const diff = daysBetween(h, today);
        const cls = diff < 0 ? 'late' : (diff <= 3 ? 'soon' : 'nx');
        const txt = diff < 0 ? `逾期${-diff}天` : (diff === 0 ? '今天到期' : `还剩${diff}天`);
        nx = `<div class="${cls}">📅 ${esc(h)}（${txt}）</div>`;
      }
      const overdue = h && daysBetween(h, today) < 0 && !FINAL.includes(r['状态']);

      return `<tr data-row="${r._row}" class="${overdue ? 'overdue' : ''}">
        <td class="days">${r.序号 ?? ''}</td>
        <td class="co">${esc(r.公司)}</td>
        <td class="pos">${esc(r.岗位)}</td>
        <td>${esc(d)}</td>
        <td>${esc(r.渠道)}</td>
        <td><select class="stsel ${statusClass(r['状态'])}" data-row="${r._row}">
              ${opts('状态', r['状态'])}
            </select></td>
        <td class="next">${esc(r['下一步动作'])}${nx ? '<br>' + nx : ''}</td>
        <td class="days">${days === '' ? '' : days + '天'}</td>
        <td><div class="ops">
          <button class="op" data-act="edit" data-row="${r._row}" type="button">编辑</button>
          <button class="op del" data-act="del" data-row="${r._row}" type="button">删除</button>
        </div></td>
      </tr>`;
    }).join('');
  }

  /* ---------------- 表单 ---------------- */
  function collect() {
    const o = {};
    ALL.forEach(n => { o[n] = (el(n).value || '').trim(); });
    return o;
  }

  function fillForm(r) {
    ALL.forEach(n => { el(n).value = r[n] || ''; });
    $('f__row').value = r._row;
    $('f__expect').value = JSON.stringify([r.公司, r.岗位]);
    editing = true;
    $('formTitle').textContent = `✏️ 正在修改：${r.公司} · ${r.岗位}`;
    $('btnSubmit').textContent = '保存修改';
    $('btnCancelEdit').classList.remove('hidden');
    window.scrollTo({ top: 0, behavior: 'smooth' });
    el('公司').focus();
  }

  function clearEdit() {
    editing = false;
    $('form').reset();
    $('f__row').value = '';
    $('f__expect').value = '';
    ALL.forEach(n => { el(n).value = ''; });
    $('formTitle').textContent = '➕ 新增投递记录';
    $('btnSubmit').textContent = '提交写入 Excel';
    $('btnCancelEdit').classList.add('hidden');
    applyDefaults();
  }

  function applyDefaults() {
    el('投递日期').value = todayStr();
    el('状态').value = '已投递';
    const lastCh = localStorage.getItem('last_channel');
    if (lastCh && (options['渠道'] || []).includes(lastCh)) el('渠道').value = lastCh;
  }

  async function submit(e) {
    e.preventDefault();
    const data = collect();
    if (!data.公司 || !data.岗位) { toast('「公司」和「岗位」必填', 'err'); return; }

    const btn = $('btnSubmit');
    btn.disabled = true;
    const old = btn.textContent;
    btn.textContent = editing ? '保存中…' : '写入中…';
    try {
      if (editing) {
        data._row = Number($('f__row').value);
        data._expect = JSON.parse($('f__expect').value || 'null');
        const r = await api('/api/update', data);
        toast('✓ ' + r.message, 'ok');
        clearEdit();
      } else {
        const r = await api('/api/add', data);
        toast('✓ ' + r.message, 'ok');
        const keepCh = data.渠道;
        $('form').reset();
        ALL.forEach(n => { el(n).value = ''; });
        applyDefaults();
        el('渠道').value = keepCh || el('渠道').value;
        el('公司').focus();
      }
      if (data.渠道) localStorage.setItem('last_channel', data.渠道);
      await load();
    } catch (err) {
      toast('✗ ' + err.message, 'err');
    } finally {
      btn.disabled = false;
      btn.textContent = editing ? '保存修改' : '提交写入 Excel';
    }
  }

  /* ---------------- 事件 ---------------- */
  $('form').addEventListener('submit', submit);
  $('btnCancelEdit').addEventListener('click', clearEdit);
  $('form').addEventListener('reset', () => setTimeout(applyDefaults, 0));
  $('btnRefresh').addEventListener('click', () => load());
  $('search').addEventListener('input', renderTable);

  $('tbody').addEventListener('change', async (e) => {
    const sel = e.target.closest('.stsel');
    if (!sel) return;
    const row = Number(sel.dataset.row);
    const rec = byRow[row];
    if (!rec) return;
    try {
      const r = await api('/api/update', {
        ...rec, 状态: sel.value, _row: row, _expect: [rec.公司, rec.岗位]
      });
      toast('✓ 状态已更新为「' + (sel.value || '空') + '」', 'ok');
      await load(true);
    } catch (err) {
      toast('✗ ' + err.message, 'err');
      await load(true);
    }
  });

  $('tbody').addEventListener('click', async (e) => {
    const b = e.target.closest('.op');
    if (!b) return;
    const row = Number(b.dataset.row);
    const rec = byRow[row];
    if (!rec) return;

    if (b.dataset.act === 'edit') { fillForm(rec); return; }

    if (b.dataset.act === 'del') {
      if (!confirm(`确定删除这条记录吗？\n\n${rec.公司} · ${rec.岗位}\n\n（可以还原：记录会先挪进「回收站」，随时能恢复）`)) return;
      try {
        const r = await api('/api/delete', { _row: row });
        toast('✓ ' + r.message, 'ok');
        if (editing && Number($('f__row').value) === row) clearEdit();
        await load();
      } catch (err) {
        toast('✗ ' + err.message, 'err');
      }
    }
  });

  /* ---- 批量导入 ---- */
  $('btnBulk').addEventListener('click', async () => {
    const text = $('bulkText').value;
    if (!text.trim()) { toast('先粘贴几行再点导入', 'err'); return; }
    const btn = $('btnBulk');
    btn.disabled = true; btn.textContent = '导入中…';
    try {
      const r = await api('/api/bulk', {
        text,
        投递日期: el('投递日期').value || todayStr(),
        渠道: el('渠道').value,
        状态: el('状态').value || '已投递',
        面试形式: el('面试形式').value,
        结果: el('结果').value,
      });
      const box = $('bulkResult');
      box.classList.remove('hidden');
      box.innerHTML =
        `<div><b>${esc(r.message)}</b></div>` +
        (r.added.length ? '<div class="ok-line">✓ ' + r.added.map(esc).join('　') + '</div>' : '') +
        (r.skipped.length ? '<div class="skip-line">跳过：' + r.skipped.map(esc).join('　') + '</div>' : '');
      toast('✓ ' + r.message, 'ok');
      $('bulkText').value = '';
      await load(true);
    } catch (err) {
      toast('✗ ' + err.message, 'err');
    } finally {
      btn.disabled = false; btn.textContent = '导入到 Excel';
    }
  });

  /* ---- 回收站 / 历史版本 ---- */
  $('btnToggleSafety').addEventListener('click', () => {
    const b = $('safetyBody');
    const open = b.classList.toggle('hidden') === false;
    $('btnToggleSafety').textContent = open ? '收起' : '展开';
  });

  $('btnBackup').addEventListener('click', async () => {
    try {
      const r = await api('/api/backup', {});
      toast('✓ ' + r.message + (r.name ? '：' + r.name : ''), 'ok');
      await load(true);
    } catch (err) { toast('✗ ' + err.message, 'err'); }
  });

  $('trashBody').addEventListener('click', async (e) => {
    const b = e.target.closest('.op');
    if (!b || b.dataset.act !== 'untrash') return;
    try {
      const r = await api('/api/trash/restore', { _trow: Number(b.dataset.trow) });
      toast('✓ ' + r.message, 'ok');
      await load(true);
    } catch (err) { toast('✗ ' + err.message, 'err'); }
  });

  $('backupList').addEventListener('click', async (e) => {
    const b = e.target.closest('.op');
    if (!b || b.dataset.act !== 'restore') return;
    const name = b.dataset.name;
    if (!confirm(`把整张表回滚到这个版本吗？\n\n${name}\n\n` +
      `当前状态会先自动备份一份，回滚错了还能再退回来。`)) return;
    try {
      const r = await api('/api/restore', { name });
      toast('✓ ' + r.message, 'ok');
      await load(true);
    } catch (err) { toast('✗ ' + err.message, 'err'); }
  });

  document.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
      e.preventDefault();
      $('form').requestSubmit();
    }
    if (e.altKey && /^[1-9]$/.test(e.key)) {
      const order = ['公司', '岗位', '投递日期', '渠道', '状态', '下一步动作', '截止/跟进时间', '备注', '笔试时间'];
      const n = order[Number(e.key) - 1];
      if (n) { e.preventDefault(); el(n).focus(); if (el(n).select) el(n).select(); }
    }
  });

  /* ---------------- 启动 ---------------- */
  applyDefaults();
  load(true).then(() => el('公司').focus());
})();
