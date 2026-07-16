import { NextRequest, NextResponse } from "next/server";

const SUPABASE_URL = process.env.NEXT_PUBLIC_SUPABASE_URL;
const SUPABASE_SECRET_KEY = process.env.SUPABASE_SECRET_KEY;

type ChecklistKey = "main" | "annex" | "restroom";
const categoryToDb: Record<ChecklistKey, string> = {
  main: "main_facility",
  annex: "annex_facility",
  restroom: "restroom",
};
const dbToCategory: Record<string, ChecklistKey> = {
  main_facility: "main",
  annex_facility: "annex",
  restroom: "restroom",
};

const MAIN_NAMES = [
  "본동 출입구 주변 청결", "현관 바닥 및 매트 청결", "1층 복도 바닥 청결", "2층 복도 바닥 청결", "계단 및 난간 청결", "엘리베이터 내부 청결", "공용 창문 및 창틀 청결", "공용 유리문 청결", "복도 조명 및 스위치 청결", "소화기 주변 정리", "냉온수기 외부 청결", "냉온수기 물받이 청결", "공용 쓰레기통 청결", "분리수거함 정리", "휴게공간 탁자 청결", "휴게공간 의자 정돈", "객실 출입문 청결", "객실 바닥 청결", "객실 침대 및 침구 정돈", "객실 탁자 청결", "객실 수납장 청결", "객실 창문 및 창틀 청결", "객실 테라스 청결", "객실 세면대 청결", "객실 거울 청결", "객실 화장실 변기 청결", "객실 화장실 바닥 청결", "객실 휴지통 청결", "객실 비품 정리", "교육장 바닥 청결", "교육장 책상 청결", "교육장 의자 정돈", "교육장 기자재 주변 청결", "강사대 및 칠판 청결", "식당 출입구 청결", "식당 바닥 청결", "공용 세면 공간 청결", "샤워실 바닥 청결", "샤워실 비품 정리", "비상구 주변 정리", "기타 이상 여부"
];
const ANNEX_NAMES = [
  "부속동 출입구 주변 청결", "현관 바닥 청결", "현관 매트 청결", "신발장 내부 정리", "복도 바닥 청결", "계단 및 난간 청결", "창문 및 창틀 청결", "유리문 청결", "공용 조명 및 스위치 청결", "냉온수기 외부 청결", "냉온수기 물받이 청결", "공용 쓰레기통 청결", "분리수거함 정리", "객실 출입문 청결", "객실 바닥 청결", "객실 침대 및 침구 정돈", "객실 탁자 청결", "객실 수납장 청결", "객실 창문 및 창틀 청결", "객실 세면대 청결", "객실 거울 청결", "객실 화장실 변기 청결", "객실 화장실 바닥 청결", "객실 휴지통 청결", "객실 비품 정리", "샤워실 바닥 청결", "샤워실 배수 상태", "샤워실 발매트 청결", "샤워실 비품 정리", "기념관 바닥 청결", "기념관 전시대 청결", "기념관 탁자 및 의자 정돈", "유르트 출입구 청결", "유르트 바닥 청결", "유르트 창문 청결", "유르트 비품 정리", "공기청정기 외부 청결", "공기청정기 작동 상태", "소화기 주변 정리", "비상구 주변 정리", "외부 데크 및 계단 청결", "기타 이상 여부"
];
const RESTROOM_NAMES = ["소변기 작동 및 청결", "좌변기 작동 및 청결", "세면대 청결", "세면대 배수 상태", "거울 청결", "바닥 청결", "조명등 상태", "쓰레기통 청결", "화장지·비누 비치", "냄새 및 악취 여부", "기타 이상 여부"];

function assertConfig() {
  if (!SUPABASE_URL || !SUPABASE_SECRET_KEY) throw new Error("Supabase 환경변수가 설정되지 않았습니다.");
}

async function db(path: string, init: RequestInit = {}) {
  assertConfig();
  const response = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    ...init,
    headers: {
      apikey: SUPABASE_SECRET_KEY!,
      Authorization: `Bearer ${SUPABASE_SECRET_KEY}`,
      "Content-Type": "application/json",
      Prefer: "return=representation",
      ...(init.headers || {}),
    },
    cache: "no-store",
  });
  if (!response.ok) {
    const text = await response.text();
    throw new Error(`${response.status} ${text}`);
  }
  const text = await response.text();
  return text ? JSON.parse(text) : null;
}

async function seedDefaults() {
  const items = await db("checklist_items?select=id&limit=1");
  if (!items.length) {
    const rows = [
      ...MAIN_NAMES.map((item_name, i) => ({ category: "main_facility", item_name, is_active: true, sort_order: i })),
      ...ANNEX_NAMES.map((item_name, i) => ({ category: "annex_facility", item_name, is_active: true, sort_order: i })),
      ...RESTROOM_NAMES.map((item_name, i) => ({ category: "restroom", item_name, is_active: true, sort_order: i })),
    ];
    await db("checklist_items", { method: "POST", body: JSON.stringify(rows) });
  }
}

async function loadAll() {
  await seedDefaults();
  const [inspectors, checklistItems, records, recordItems, revisions, settings] = await Promise.all([
    db("inspectors?select=*&order=sort_order.asc,created_at.asc"),
    db("checklist_items?select=*&order=category.asc,sort_order.asc,created_at.asc"),
    db("inspection_records?select=*&order=inspection_date.desc,created_at.desc"),
    db("inspection_record_items?select=*&order=sort_order.asc,created_at.asc"),
    db("inspection_revisions?select=*&order=created_at.asc"),
    db("app_settings?select=setting_key,setting_value"),
  ]);

  const staff = inspectors.map((row: any) => ({ id: row.id, name: row.name, active: row.is_active }));
  const managedItems: Record<ChecklistKey, any[]> = { main: [], annex: [], restroom: [] };
  checklistItems.forEach((row: any) => {
    const key = dbToCategory[row.category];
    if (key) managedItems[key].push({ id: row.id, name: row.item_name, active: row.is_active });
  });
  const itemsByRecord = new Map<string, any[]>();
  recordItems.forEach((row: any) => {
    const list = itemsByRecord.get(row.record_id) || [];
    list.push({ itemId: row.checklist_item_id || row.id, name: row.item_name, result: row.result, note: row.note || "" });
    itemsByRecord.set(row.record_id, list);
  });
  const revisionsByRecord = new Map<string, any[]>();
  revisions.forEach((row: any) => {
    const list = revisionsByRecord.get(row.record_id) || [];
    list.push({ at: new Date(row.created_at).toLocaleString("ko-KR", { dateStyle: "medium", timeStyle: "short" }), summary: row.revision_summary });
    revisionsByRecord.set(row.record_id, list);
  });
  const mappedRecords = records.map((row: any) => ({
    id: row.id,
    building: row.building,
    kind: row.inspection_kind,
    date: row.inspection_date,
    inspectors: [row.inspector_1_name, row.inspector_2_name],
    createdAt: new Date(row.created_at).toLocaleString("ko-KR", { dateStyle: "medium", timeStyle: "short" }),
    updatedAt: new Date(row.updated_at).toLocaleString("ko-KR", { dateStyle: "medium", timeStyle: "short" }),
    items: itemsByRecord.get(row.id) || [],
    revisions: revisionsByRecord.get(row.id) || [],
  }));
  const settingMap = Object.fromEntries(settings.map((row: any) => [row.setting_key, row.setting_value]));
  let excludedDates = [];
  try { excludedDates = JSON.parse(settingMap.excluded_dates || "[]"); } catch { excludedDates = []; }
  return { staff, managedItems, records: mappedRecords, adminPassword: settingMap.admin_password || "482915", excludedDates };
}

export async function GET() {
  try {
    return NextResponse.json(await loadAll());
  } catch (error) {
    return NextResponse.json({ error: error instanceof Error ? error.message : "데이터를 불러오지 못했습니다." }, { status: 500 });
  }
}

export async function POST(request: NextRequest) {
  try {
    const body = await request.json();
    const { action } = body;

    if (action === "saveRecord") {
      const r = body.record;
      let saved;
      if (r.id && /^[0-9a-f-]{36}$/i.test(r.id)) {
        saved = await db(`inspection_records?id=eq.${r.id}`, { method: "PATCH", body: JSON.stringify({ inspector_1_id: null, inspector_2_id: null, inspector_1_name: r.inspectors[0], inspector_2_name: r.inspectors[1], updated_at: new Date().toISOString() }) });
      } else {
        saved = await db("inspection_records?on_conflict=inspection_date,building,inspection_kind", {
          method: "POST",
          headers: { Prefer: "resolution=merge-duplicates,return=representation" },
          body: JSON.stringify({ inspection_date: r.date, building: r.building, inspection_kind: r.kind, inspector_1_id: null, inspector_2_id: null, inspector_1_name: r.inspectors[0], inspector_2_name: r.inspectors[1], updated_at: new Date().toISOString() }),
        });
      }
      const recordId = saved[0].id;
      await db(`inspection_record_items?record_id=eq.${recordId}`, { method: "DELETE" });
      if (r.items.length) {
        await db("inspection_record_items", { method: "POST", body: JSON.stringify(r.items.map((item: any, index: number) => ({ record_id: recordId, checklist_item_id: /^[0-9a-f-]{36}$/i.test(item.itemId) ? item.itemId : null, item_name: item.name, result: item.result, note: item.note || "", sort_order: index }))) });
      }
      if (body.revisionSummary) {
        await db("inspection_revisions", { method: "POST", body: JSON.stringify({ record_id: recordId, revision_summary: body.revisionSummary, revised_by: r.inspectors.join(" · ") }) });
      }
      return NextResponse.json(await loadAll());
    }

    if (action === "addStaff") {
      await db("inspectors", { method: "POST", body: JSON.stringify({ name: body.name, is_active: true, sort_order: body.sortOrder ?? 0 }) });
    } else if (action === "updateStaff") {
      await db(`inspectors?id=eq.${body.id}`, { method: "PATCH", body: JSON.stringify({ name: body.name, updated_at: new Date().toISOString() }) });
    } else if (action === "toggleStaff") {
      await db(`inspectors?id=eq.${body.id}`, { method: "PATCH", body: JSON.stringify({ is_active: body.active, updated_at: new Date().toISOString() }) });
    } else if (action === "deleteStaff") {
      await db(`inspectors?id=eq.${body.id}`, { method: "DELETE" });
    } else if (action === "addItem") {
      await db("checklist_items", { method: "POST", body: JSON.stringify({ category: categoryToDb[body.category as ChecklistKey], item_name: body.name, is_active: true, sort_order: body.sortOrder ?? 0 }) });
    } else if (action === "updateItem") {
      await db(`checklist_items?id=eq.${body.id}`, { method: "PATCH", body: JSON.stringify({ item_name: body.name, updated_at: new Date().toISOString() }) });
    } else if (action === "toggleItem") {
      await db(`checklist_items?id=eq.${body.id}`, { method: "PATCH", body: JSON.stringify({ is_active: body.active, updated_at: new Date().toISOString() }) });
    } else if (action === "deleteItem") {
      await db(`checklist_items?id=eq.${body.id}`, { method: "DELETE" });
    } else if (action === "reorderItems") {
      await Promise.all(body.ids.map((id: string, index: number) => db(`checklist_items?id=eq.${id}`, { method: "PATCH", body: JSON.stringify({ sort_order: index, updated_at: new Date().toISOString() }) })));
    } else if (action === "updateExcludedDates") {
      const value = JSON.stringify(Array.isArray(body.excludedDates) ? body.excludedDates : []);
      await db("app_settings?on_conflict=setting_key", { method: "POST", headers: { Prefer: "resolution=merge-duplicates,return=representation" }, body: JSON.stringify({ setting_key: "excluded_dates", setting_value: value, updated_at: new Date().toISOString() }) });
    } else if (action === "updateAdminPassword") {
      await db("app_settings?on_conflict=setting_key", { method: "POST", headers: { Prefer: "resolution=merge-duplicates,return=representation" }, body: JSON.stringify({ setting_key: "admin_password", setting_value: body.password, updated_at: new Date().toISOString() }) });
    } else {
      return NextResponse.json({ error: "지원하지 않는 작업입니다." }, { status: 400 });
    }
    return NextResponse.json(await loadAll());
  } catch (error) {
    return NextResponse.json({ error: error instanceof Error ? error.message : "저장하지 못했습니다." }, { status: 500 });
  }
}
