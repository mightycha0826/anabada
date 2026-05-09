import { supabaseAdmin } from '$lib/supabaseAdmin.js';
import { ADMIN_PASSWORD } from '$env/static/private';
import * as XLSX from 'xlsx';

const SESSION_COOKIE = 'admin_session';

const HEADERS = ['제목', '가격 (비즈쿨 머니)', '영역', '반', '모둠', '예약자 학번', '예약자 이메일'];
const DOMAIN_LABELS = { 1: '1영역 (책·학습)', 2: '2영역 (의류·액세서리)', 3: '3영역 (취미·굿즈)' };

function itemToRow(item) {
  return [
    item.title,
    item.price ?? 0,
    item.domain ? DOMAIN_LABELS[item.domain] : '',
    `${item.class_num}반`,
    `${item.group_num}모둠`,
    item.reserved_by ?? '',
    item.user_email ?? ''
  ];
}

function buildSheet(items) {
  const rows = [HEADERS, ...items.map(itemToRow)];
  const total = items.reduce((s, i) => s + (i.price ?? 0), 0);
  rows.push(['합계', total, '', '', '', '', '']);
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [{ wch: 32 }, { wch: 18 }, { wch: 22 }, { wch: 8 }, { wch: 8 }, { wch: 14 }, { wch: 28 }];
  return ws;
}

export async function GET({ url, cookies }) {
  if (cookies.get(SESSION_COOKIE) !== ADMIN_PASSWORD) {
    return new Response('Unauthorized', { status: 401 });
  }

  const type = url.searchParams.get('type') ?? 'all';

  const { data: items, error } = await supabaseAdmin
    .from('items')
    .select('*')
    .order('class_num')
    .order('group_num')
    .order('created_at', { ascending: false });

  if (error) return new Response('DB error', { status: 500 });

  const all = items ?? [];
  const wb = XLSX.utils.book_new();

  if (type === 'class') {
    for (let c = 1; c <= 12; c++) {
      const subset = all.filter((i) => i.class_num === c);
      if (subset.length > 0) {
        XLSX.utils.book_append_sheet(wb, buildSheet(subset), `${c}반`);
      }
    }
    if (wb.SheetNames.length === 0) {
      XLSX.utils.book_append_sheet(wb, buildSheet([]), '데이터 없음');
    }
  } else if (type === 'domain') {
    for (let d = 1; d <= 3; d++) {
      const subset = all.filter((i) => i.domain === d);
      if (subset.length > 0) {
        XLSX.utils.book_append_sheet(wb, buildSheet(subset), DOMAIN_LABELS[d]);
      }
    }
    if (wb.SheetNames.length === 0) {
      XLSX.utils.book_append_sheet(wb, buildSheet([]), '데이터 없음');
    }
  } else {
    XLSX.utils.book_append_sheet(wb, buildSheet(all), '전체 물품');
  }

  const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  const label = type === 'class' ? '반별' : type === 'domain' ? '영역별' : '전체';
  const date  = new Date().toISOString().slice(0, 10);
  const filename = encodeURIComponent(`아나바다_${label}_${date}.xlsx`);

  return new Response(buf, {
    headers: {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename*=UTF-8''${filename}`
    }
  });
}
