"use client";

import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

type Building = "본동" | "부속동";
type InspectionKind = "시설청소" | "남WC" | "여WC";
type Result = "정상" | "이상";
type Screen = "calendar" | "inspectors" | "checklist" | "detail" | "restroomMonthly" | "facilityMonthly" | "adminGate" | "admin";
type AdminTab = "staff" | "items" | "closedDates" | "export" | "security";
type ChecklistKey = "main" | "annex" | "restroom";

type StaffMember = { id: string; name: string; active: boolean };
type ManagedItem = { id: string; name: string; active: boolean };
type Revision = { at: string; summary: string };
type RecordItem = { itemId: string; name: string; result: Result; note: string };
type InspectionRecord = {
  id: string;
  building: Building;
  kind: InspectionKind;
  date: string;
  inspectors: [string, string];
  createdAt: string;
  updatedAt: string;
  items: RecordItem[];
  revisions: Revision[];
};

type Status = "normal" | "issue" | "missing";
type ExcludedDate = { date: string; reason: string };



const MAIN_NAMES = [
  "본동 출입구 주변 청결", "현관 바닥 및 매트 청결", "1층 복도 바닥 청결", "2층 복도 바닥 청결",
  "계단 및 난간 청결", "엘리베이터 내부 청결", "공용 창문 및 창틀 청결", "공용 유리문 청결",
  "복도 조명 및 스위치 청결", "소화기 주변 정리", "냉온수기 외부 청결", "냉온수기 물받이 청결",
  "공용 쓰레기통 청결", "분리수거함 정리", "휴게공간 탁자 청결", "휴게공간 의자 정돈",
  "객실 출입문 청결", "객실 바닥 청결", "객실 침대 및 침구 정돈", "객실 탁자 청결",
  "객실 수납장 청결", "객실 창문 및 창틀 청결", "객실 테라스 청결", "객실 세면대 청결",
  "객실 거울 청결", "객실 화장실 변기 청결", "객실 화장실 바닥 청결", "객실 휴지통 청결",
  "객실 비품 정리", "교육장 바닥 청결", "교육장 책상 청결", "교육장 의자 정돈",
  "교육장 기자재 주변 청결", "강사대 및 칠판 청결", "식당 출입구 청결", "식당 바닥 청결",
  "공용 세면 공간 청결", "샤워실 바닥 청결", "샤워실 비품 정리", "비상구 주변 정리", "기타 이상 여부"
];

const ANNEX_NAMES = [
  "부속동 출입구 주변 청결", "현관 바닥 청결", "현관 매트 청결", "신발장 내부 정리",
  "복도 바닥 청결", "계단 및 난간 청결", "창문 및 창틀 청결", "유리문 청결",
  "공용 조명 및 스위치 청결", "냉온수기 외부 청결", "냉온수기 물받이 청결", "공용 쓰레기통 청결",
  "분리수거함 정리", "객실 출입문 청결", "객실 바닥 청결", "객실 침대 및 침구 정돈",
  "객실 탁자 청결", "객실 수납장 청결", "객실 창문 및 창틀 청결", "객실 세면대 청결",
  "객실 거울 청결", "객실 화장실 변기 청결", "객실 화장실 바닥 청결", "객실 휴지통 청결",
  "객실 비품 정리", "샤워실 바닥 청결", "샤워실 배수 상태", "샤워실 발매트 청결",
  "샤워실 비품 정리", "기념관 바닥 청결", "기념관 전시대 청결", "기념관 탁자 및 의자 정돈",
  "유르트 출입구 청결", "유르트 바닥 청결", "유르트 창문 청결", "유르트 비품 정리",
  "공기청정기 외부 청결", "공기청정기 작동 상태", "소화기 주변 정리", "비상구 주변 정리",
  "외부 데크 및 계단 청결", "기타 이상 여부"
];

const RESTROOM_NAMES = [
  "소변기 작동 및 청결", "좌변기 작동 및 청결", "세면대 청결", "세면대 배수 상태", "거울 청결",
  "바닥 청결", "조명등 상태", "쓰레기통 청결", "화장지·비누 비치", "냄새 및 악취 여부", "기타 이상 여부"
];

function makeItems(prefix: string, names: string[]): ManagedItem[] {
  return names.map((name, index) => ({ id: `${prefix}-${index + 1}`, name, active: true }));
}

const DEFAULT_ITEMS: Record<ChecklistKey, ManagedItem[]> = {
  main: makeItems("main", MAIN_NAMES),
  annex: makeItems("annex", ANNEX_NAMES),
  restroom: makeItems("restroom", RESTROOM_NAMES)
};

// 2026 public holidays: 2026 calendar standard plus public holidays newly added in 2026.
const HOLIDAYS_2026: Record<string, string> = {
  "2026-01-01": "신정",
  "2026-02-16": "설날 연휴", "2026-02-17": "설날", "2026-02-18": "설날 연휴",
  "2026-03-01": "삼일절", "2026-03-02": "대체공휴일",
  "2026-05-01": "노동절", "2026-05-05": "어린이날", "2026-05-24": "부처님오신날", "2026-05-25": "대체공휴일",
  "2026-06-03": "지방선거일", "2026-06-06": "현충일",
  "2026-07-17": "제헌절",
  "2026-08-15": "광복절", "2026-08-17": "대체공휴일",
  "2026-09-24": "추석 연휴", "2026-09-25": "추석", "2026-09-26": "추석 연휴",
  "2026-10-03": "개천절", "2026-10-05": "대체공휴일", "2026-10-09": "한글날",
  "2026-12-25": "기독탄신일"
};

const LS = {
  records: "pa-hygiene-v8-records",
  staff: "pa-hygiene-v8-staff",
  items: "pa-hygiene-v8-items",
  recent: "pa-hygiene-v8-recent-inspectors",
  adminPassword: "pa-hygiene-v9-admin-password"
};

function toKey(date = new Date()) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}
function parseKey(key: string) { return new Date(`${key}T12:00:00`); }
function formatDay(key: string) { const d = parseKey(key); return `${d.getMonth() + 1}월 ${d.getDate()}일`; }
function formatShort(key: string) { const d = parseKey(key); return `${d.getMonth() + 1}/${d.getDate()}`; }
function formatMonth(date: Date) { return `${date.getFullYear()}년 ${date.getMonth() + 1}월`; }
function isWeekend(key: string) { const day = parseKey(key).getDay(); return day === 0 || day === 6; }
function holidayName(key: string) { return HOLIDAYS_2026[key] ?? ""; }
function isClosed(key: string) { return isWeekend(key) || Boolean(holidayName(key)); }
function isFuture(key: string) { return key > toKey(); }
function mondayKey(key: string) {
  const d = parseKey(key); const day = d.getDay(); const diff = day === 0 ? -6 : 1 - day;
  d.setDate(d.getDate() + diff); return toKey(d);
}
function fridayKey(key: string) { const d = parseKey(mondayKey(key)); d.setDate(d.getDate() + 4); return toKey(d); }
function weekLabel(key: string) { return `${formatShort(mondayKey(key))} ~ ${formatShort(fridayKey(key))}`; }
function calendarCells(date: Date) {
  const year = date.getFullYear(); const month = date.getMonth();
  const first = new Date(year, month, 1).getDay(); const last = new Date(year, month + 1, 0).getDate();
  const cells: Array<number | null> = Array(first).fill(null);
  for (let day = 1; day <= last; day += 1) cells.push(day);
  while (cells.length % 7) cells.push(null);
  return cells;
}
function monthKey(date: Date) { return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`; }
function restroomLabel(kind: "남WC" | "여WC") { return kind === "남WC" ? "남자 화장실" : "여자 화장실"; }
function kindLabel(kind: InspectionKind) { return kind === "시설청소" ? "시설 청소" : restroomLabel(kind); }
function recordStatus(record?: InspectionRecord): Status {
  if (!record) return "missing";
  return record.items.some((item) => item.result === "이상") ? "issue" : "normal";
}
function statusLabel(status: Status) {
  if (status === "normal") return "완료·이상없음";
  if (status === "issue") return "완료·이상있음";
  return "미점검";
}
function checklistKey(building: Building, kind: InspectionKind): ChecklistKey {
  if (kind !== "시설청소") return "restroom";
  return building === "본동" ? "main" : "annex";
}
function uniqueId(prefix: string) { return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`; }
function weeksForMonth(date: Date) {
  const year = date.getFullYear(); const month = date.getMonth(); const last = new Date(year, month + 1, 0).getDate();
  const weeks: string[] = [];
  for (let day = 1; day <= last; day += 1) {
    const key = toKey(new Date(year, month, day)); const mon = mondayKey(key);
    if (!weeks.includes(mon)) weeks.push(mon);
  }
  return weeks;
}

export default function Page() {
  const [screen, setScreen] = useState<Screen>("calendar");
  const [month, setMonth] = useState(new Date());
  const [selectedDate, setSelectedDate] = useState(toKey());
  const [records, setRecords] = useState<InspectionRecord[]>([]);
  const [staff, setStaff] = useState<StaffMember[]>([]);
  const [managedItems, setManagedItems] = useState<Record<ChecklistKey, ManagedItem[]>>(DEFAULT_ITEMS);
  const [recentInspectors, setRecentInspectors] = useState<[string, string]>(["", ""]);
  const [targetBuilding, setTargetBuilding] = useState<Building>("본동");
  const [targetKind, setTargetKind] = useState<InspectionKind>("시설청소");
  const [targetRecordId, setTargetRecordId] = useState<string | null>(null);
  const [inspector1, setInspector1] = useState("");
  const [inspector2, setInspector2] = useState("");
  const [sessionInspectors, setSessionInspectors] = useState<[string, string] | null>(null);
  const [answers, setAnswers] = useState<Record<string, Result>>({});
  const [notes, setNotes] = useState<Record<string, string>>({});
  const [editing, setEditing] = useState(false);
  const [adminTab, setAdminTab] = useState<AdminTab>("staff");
  const [itemTab, setItemTab] = useState<ChecklistKey>("main");
  const [exportMonth, setExportMonth] = useState(monthKey(new Date()));
  const [toast, setToast] = useState("");
  const [hydrated, setHydrated] = useState(false);
  const [adminAuthorized, setAdminAuthorized] = useState(false);
  const [adminPassword, setAdminPassword] = useState("482915");
  const [excludedDates, setExcludedDates] = useState<ExcludedDate[]>([]);
  const [excludedDateDraft, setExcludedDateDraft] = useState("");
  const [excludedReasonDraft, setExcludedReasonDraft] = useState("");
  const [excludedEditingDate, setExcludedEditingDate] = useState<string | null>(null);
  const [adminInput, setAdminInput] = useState("");
  const [securityCurrent, setSecurityCurrent] = useState("");
  const [securityNext, setSecurityNext] = useState("");
  const [securityConfirm, setSecurityConfirm] = useState("");
  const [staffEditId, setStaffEditId] = useState<string | null>(null);
  const [staffDraft, setStaffDraft] = useState("");
  const [staffAdding, setStaffAdding] = useState(false);
  const [itemEditId, setItemEditId] = useState<string | null>(null);
  const [itemDraft, setItemDraft] = useState("");
  const [itemAdding, setItemAdding] = useState(false);
  const [dataLoading, setDataLoading] = useState(true);
  const [dataSaving, setDataSaving] = useState(false);
  const [dataError, setDataError] = useState("");

  useEffect(() => {
    let cancelled = false;
    async function loadSharedData() {
      try {
        const response = await fetch("/api/data", { cache: "no-store" });
        const payload = await response.json();
        if (!response.ok) throw new Error(payload.error || "공용 데이터를 불러오지 못했습니다.");
        if (cancelled) return;
        setRecords(payload.records);
        setStaff(payload.staff);
        setManagedItems(payload.managedItems);
        setAdminPassword(payload.adminPassword || "482915");
        setExcludedDates(payload.excludedDates || []);
        const savedRecent = localStorage.getItem(LS.recent);
        if (savedRecent) {
          try {
            const parsed = JSON.parse(savedRecent) as [string, string];
            const activeNames = new Set(
              payload.staff
                .filter((person: StaffMember) => person.active)
                .map((person: StaffMember) => person.name)
            );
            if (activeNames.has(parsed[0]) && activeNames.has(parsed[1])) {
              setRecentInspectors(parsed);
            } else {
              localStorage.removeItem(LS.recent);
              setRecentInspectors(["", ""]);
            }
          } catch {
            localStorage.removeItem(LS.recent);
            setRecentInspectors(["", ""]);
          }
        }
        setDataError("");
      } catch (error) {
        if (!cancelled) setDataError(error instanceof Error ? error.message : "공용 데이터를 불러오지 못했습니다.");
      } finally {
        if (!cancelled) { setHydrated(true); setDataLoading(false); }
      }
    }
    loadSharedData();
    const timer = window.setInterval(loadSharedData, 30000);
    return () => { cancelled = true; window.clearInterval(timer); };
  }, []);

  useEffect(() => { if (hydrated) localStorage.setItem(LS.recent, JSON.stringify(recentInspectors)); }, [recentInspectors, hydrated]);

  async function apiAction(body: Record<string, unknown>) {
    setDataSaving(true);
    try {
      const response = await fetch("/api/data", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
      const payload = await response.json();
      if (!response.ok) throw new Error(payload.error || "공용 데이터 저장에 실패했습니다.");
      setRecords(payload.records);
      setStaff(payload.staff);
      setManagedItems(payload.managedItems);
      setAdminPassword(payload.adminPassword || "482915");
      setExcludedDates(payload.excludedDates || []);
      setDataError("");
      return payload;
    } catch (error) {
      const message = error instanceof Error ? error.message : "공용 데이터 저장에 실패했습니다.";
      setDataError(message);
      throw error;
    } finally {
      setDataSaving(false);
    }
  }


  function excludedReason(key: string) {
    return excludedDates.find((item) => item.date === key)?.reason || "";
  }

  function isInspectionClosed(key: string) {
    return isClosed(key) || Boolean(excludedReason(key));
  }

  function isFacilityWeekExcluded(week: string) {
    const monday = parseKey(mondayKey(week));
    return Array.from({ length: 5 }, (_, index) => {
      const date = new Date(monday);
      date.setDate(monday.getDate() + index);
      return isInspectionClosed(toKey(date));
    }).every(Boolean);
  }

  function resetExcludedDateEditor() {
    setExcludedDateDraft("");
    setExcludedReasonDraft("");
    setExcludedEditingDate(null);
  }

  function startEditExcludedDate(item: ExcludedDate) {
    setExcludedEditingDate(item.date);
    setExcludedDateDraft(item.date);
    setExcludedReasonDraft(item.reason);
  }

  async function saveExcludedDate() {
    const date = excludedDateDraft.trim();
    const reason = excludedReasonDraft.trim();
    if (!date || !reason) return flash("날짜와 제외 사유를 입력해 주세요.");
    const next = excludedDates
      .filter((item) => item.date !== excludedEditingDate && item.date !== date)
      .concat({ date, reason })
      .sort((a, b) => a.date.localeCompare(b.date));
    try { await apiAction({ action: "updateExcludedDates", excludedDates: next }); } catch { return; }
    resetExcludedDateEditor();
    flash("점검 제외일을 저장했습니다.");
  }

  async function deleteExcludedDate(date: string) {
    const next = excludedDates.filter((item) => item.date !== date);
    try { await apiAction({ action: "updateExcludedDates", excludedDates: next }); } catch { return; }
    if (excludedEditingDate === date) resetExcludedDateEditor();
    flash("점검 제외일을 삭제했습니다.");
  }

  const calendarDays = useMemo(() => calendarCells(month), [month]);
  const activeStaff = staff.filter((person) => person.active);
  const activeChecklist = managedItems[checklistKey(targetBuilding, targetKind)].filter((item) => item.active);
  const selectedWeekStart = mondayKey(selectedDate);
  const selectedWeekEnd = fridayKey(selectedDate);
  const currentRecord = targetRecordId ? records.find((record) => record.id === targetRecordId) : undefined;

  function flash(text: string) {
    setToast(text);
    window.setTimeout(() => setToast(""), 2200);
  }

  function openAdmin() {
    setScreen(adminAuthorized ? "admin" : "adminGate");
  }

  function submitAdminAccess() {
    if (!adminInput.trim()) {
      flash("관리자 비밀번호를 입력해 주세요.");
      return;
    }
    if (adminInput !== adminPassword) {
      flash("비밀번호가 일치하지 않습니다.");
      return;
    }
    setAdminAuthorized(true);
    setAdminInput("");
    setScreen("admin");
  }

  function logoutAdmin() {
    setAdminAuthorized(false);
    setAdminInput("");
    setScreen("calendar");
  }

  async function updateAdminPassword() {
    if (securityCurrent !== adminPassword) {
      flash("현재 비밀번호가 올바르지 않습니다.");
      return;
    }
    if (securityNext.length < 4) {
      flash("새 비밀번호는 4자리 이상으로 입력해 주세요.");
      return;
    }
    if (securityNext !== securityConfirm) {
      flash("새 비밀번호 확인이 일치하지 않습니다.");
      return;
    }
    try {
      await apiAction({ action: "updateAdminPassword", password: securityNext });
    } catch { return; }
    setSecurityCurrent("");
    setSecurityNext("");
    setSecurityConfirm("");
    flash("관리자 비밀번호가 변경되었습니다.");
  }

  function dailyRecord(date: string, building: Building, kind: "남WC" | "여WC") {
    return records.find((record) => record.date === date && record.building === building && record.kind === kind);
  }

  function weeklyRecord(date: string, building: Building) {
    const start = mondayKey(date); const end = fridayKey(date);
    return records.find((record) => record.kind === "시설청소" && record.building === building && record.date >= start && record.date <= end);
  }

  function openTarget(building: Building, kind: InspectionKind, date = selectedDate) {
    const existing = kind === "시설청소" ? weeklyRecord(date, building) : dailyRecord(date, building, kind);
    setTargetBuilding(building); setTargetKind(kind); setTargetRecordId(existing?.id ?? null);
    if (existing) {
      setInspector1(existing.inspectors[0]); setInspector2(existing.inspectors[1]);
      setSessionInspectors(existing.inspectors);
      const nextAnswers: Record<string, Result> = {}; const nextNotes: Record<string, string> = {};
      existing.items.forEach((item) => { nextAnswers[item.itemId] = item.result; nextNotes[item.itemId] = item.note; });
      setAnswers(nextAnswers); setNotes(nextNotes); setEditing(false); setScreen("detail");
      return;
    }
    if (isInspectionClosed(date) || isFuture(date)) return;
    setSelectedDate(date);
    setInspector1(recentInspectors[0]); setInspector2(recentInspectors[1]);
    setSessionInspectors(null);
    setAnswers({}); setNotes({}); setEditing(true); setScreen("inspectors");
  }

  function startChecklist() {
    if (!inspector1 || !inspector2) return flash("점검자 2명을 선택해 주세요.");
    if (inspector1 === inspector2) return flash("서로 다른 점검자를 선택해 주세요.");
    const selected: [string, string] = [inspector1, inspector2];
    setSessionInspectors(selected);
    setRecentInspectors(selected);
    setScreen("checklist");
  }

  function markAllNormal() {
    const next: Record<string, Result> = {};
    activeChecklist.forEach((item) => { next[item.id] = "정상"; });
    setAnswers(next);
    flash("전체 항목을 정상으로 입력했습니다.");
  }

  async function saveChecklist() {
    const effectiveInspectors: [string, string] = sessionInspectors ?? [inspector1, inspector2];
    if (!effectiveInspectors[0] || !effectiveInspectors[1] || effectiveInspectors[0] === effectiveInspectors[1]) return flash("점검자 2명을 정확히 선택해 주세요.");
    const missing = activeChecklist.filter((item) => !answers[item.id]);
    if (missing.length) {
      document.getElementById(`item-${missing[0].id}`)?.scrollIntoView({ behavior: "smooth", block: "center" });
      return flash(`미입력 항목 ${missing.length}개가 있습니다.`);
    }
    const now = new Date().toLocaleString("ko-KR", { dateStyle: "medium", timeStyle: "short" });
    const savedItems: RecordItem[] = activeChecklist.map((item) => ({
      itemId: item.id, name: item.name, result: answers[item.id], note: notes[item.id] ?? ""
    }));
    const record: InspectionRecord = {
      id: targetRecordId ?? uniqueId("record"), building: targetBuilding, kind: targetKind, date: selectedDate,
      inspectors: effectiveInspectors, createdAt: currentRecord?.createdAt ?? now, updatedAt: now, items: savedItems,
      revisions: currentRecord?.revisions ?? []
    };
    try {
      const payload = await apiAction({ action: "saveRecord", record, revisionSummary: targetRecordId ? "점검 결과 또는 점검자 수정" : "" });
      const saved = payload.records.find((item: InspectionRecord) => item.date === selectedDate && item.building === targetBuilding && item.kind === targetKind);
      if (saved) setTargetRecordId(saved.id);
    } catch { return; }
    setRecentInspectors(effectiveInspectors); setSessionInspectors(effectiveInspectors);
    setEditing(false); setScreen(targetRecordId ? "detail" : "calendar"); flash(targetRecordId ? "수정 내용을 저장했습니다." : "점검을 완료했습니다.");
  }

  function beginEdit() {
    if (!currentRecord) return;
    setInspector1(currentRecord.inspectors[0]); setInspector2(currentRecord.inspectors[1]);
    setSessionInspectors(currentRecord.inspectors);
    const nextAnswers: Record<string, Result> = {}; const nextNotes: Record<string, string> = {};
    currentRecord.items.forEach((item) => { nextAnswers[item.itemId] = item.result; nextNotes[item.itemId] = item.note; });
    setAnswers(nextAnswers); setNotes(nextNotes); setEditing(true); setScreen("checklist");
  }

  function openRestroomMonthly(building: Building, kind: "남WC" | "여WC") {
    setTargetBuilding(building); setTargetKind(kind); setScreen("restroomMonthly");
  }

  function openFacilityMonthly() { setScreen("facilityMonthly"); }

  function cancelStaffEditor() {
    setStaffEditId(null);
    setStaffDraft("");
    setStaffAdding(false);
  }

  function startAddStaff() {
    setStaffEditId(null);
    setStaffDraft("");
    setStaffAdding(true);
  }

  function startEditStaff(person: StaffMember) {
    setStaffAdding(false);
    setStaffEditId(person.id);
    setStaffDraft(person.name);
  }

  async function saveStaffEditor() {
    const name = staffDraft.trim();
    if (!name) return;
    if (staff.some((person) => person.name === name && person.id !== staffEditId)) return;
    try {
      if (staffAdding) await apiAction({ action: "addStaff", name, sortOrder: staff.length });
      else if (staffEditId) await apiAction({ action: "updateStaff", id: staffEditId, name });
    } catch { return; }
    cancelStaffEditor();
  }

  async function deleteStaff(person: StaffMember) {
    if (staffEditId === person.id) cancelStaffEditor();
    try { await apiAction({ action: "deleteStaff", id: person.id }); } catch { /* keep current list */ }
  }

  async function toggleStaff(person: StaffMember) {
    try { await apiAction({ action: "toggleStaff", id: person.id, active: !person.active }); } catch { /* keep current list */ }
  }

  function cancelItemEditor() {
    setItemEditId(null);
    setItemDraft("");
    setItemAdding(false);
  }

  function startAddChecklistItem() {
    setItemEditId(null);
    setItemDraft("");
    setItemAdding(true);
  }

  function startEditChecklistItem(item: ManagedItem) {
    setItemAdding(false);
    setItemEditId(item.id);
    setItemDraft(item.name);
  }

  async function saveChecklistItemEditor() {
    const name = itemDraft.trim();
    if (!name) return;
    if (managedItems[itemTab].some((item) => item.name === name && item.id !== itemEditId)) return;
    try {
      if (itemAdding) await apiAction({ action: "addItem", category: itemTab, name, sortOrder: managedItems[itemTab].length });
      else if (itemEditId) await apiAction({ action: "updateItem", id: itemEditId, name });
    } catch { return; }
    cancelItemEditor();
  }

  async function deleteChecklistItem(item: ManagedItem) {
    if (itemEditId === item.id) cancelItemEditor();
    try { await apiAction({ action: "deleteItem", id: item.id }); } catch { /* keep current list */ }
  }

  async function toggleChecklistItem(item: ManagedItem) {
    try { await apiAction({ action: "toggleItem", id: item.id, active: !item.active }); } catch { /* keep current list */ }
  }

  async function moveChecklistItem(index: number, direction: -1 | 1) {
    const next = [...managedItems[itemTab]]; const to = index + direction;
    if (to < 0 || to >= next.length) return;
    [next[index], next[to]] = [next[to], next[index]];
    try { await apiAction({ action: "reorderItems", ids: next.map((item) => item.id) }); } catch { /* keep current list */ }
  }

  function exportMonthlyWorkbook() {
    const [yearText, monthText] = exportMonth.split("-"); const year = Number(yearText); const m = Number(monthText);
    const days = new Date(year, m, 0).getDate(); const wb = XLSX.utils.book_new();

    function appendFacilitySheet(building: Building, sheetName: string, key: ChecklistKey) {
      const monthDate = new Date(year, m - 1, 1); const weeks = weeksForMonth(monthDate);
      const monthRecords = weeks.map((week) => weeklyRecord(week, building));
      const itemNames = managedItems[key].map((item) => item.name);
      const aoa: (string | number)[][] = [
        [`${year}년 ${m}월 ${building} 시설 청소 점검`],
        ["점검항목", ...weeks.map((week, idx) => `${idx + 1}주차\n${weekLabel(week)}`)],
        ["점검일", ...monthRecords.map((record) => record?.date ?? "-")],
        ["점검자", ...monthRecords.map((record) => record ? record.inspectors.join(" · ") : "-")],
        ...itemNames.map((name) => [name, ...monthRecords.map((record) => {
          const item = record?.items.find((target) => target.name === name);
          return item ? (item.result === "정상" ? "O" : "X") : "-";
        })])
      ];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), sheetName);
    }

    function appendRestroomSheet(building: Building, kind: "남WC" | "여WC", sheetName: string) {
      const keys = Array.from({ length: days }, (_, index) => `${yearText}-${monthText}-${String(index + 1).padStart(2, "0")}`);
      const aoa: (string | number)[][] = [
        [`${year}년 ${m}월 ${building} ${restroomLabel(kind)} 점검`],
        ["점검항목", ...keys.map((key) => parseKey(key).getDate())],
        ["점검자", ...keys.map((key) => dailyRecord(key, building, kind)?.inspectors.join(" · ") ?? "-")],
        ...managedItems.restroom.map((managed) => [managed.name, ...keys.map((key) => {
          if (isInspectionClosed(key)) return excludedReason(key) || "휴무";
          const item = dailyRecord(key, building, kind)?.items.find((target) => target.name === managed.name);
          return item ? (item.result === "정상" ? "O" : "X") : "-";
        })])
      ];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), sheetName);
    }

    appendFacilitySheet("본동", "01_본동시설", "main");
    appendFacilitySheet("부속동", "02_부속동시설", "annex");
    appendRestroomSheet("본동", "남WC", "03_본동_남자화장실");
    appendRestroomSheet("본동", "여WC", "04_본동_여자화장실");
    appendRestroomSheet("부속동", "남WC", "05_부속동_남자화장실");
    appendRestroomSheet("부속동", "여WC", "06_부속동_여자화장실");

    const abnormalRows = records
      .filter((record) => record.date.startsWith(exportMonth))
      .flatMap((record) => record.items.filter((item) => item.result === "이상").map((item) => ({
        날짜: record.date, 장소: record.building, 점검구분: kindLabel(record.kind), 점검항목: item.name,
        이상내용: item.note, 점검자1: record.inspectors[0], 점검자2: record.inspectors[1], 최종수정: record.updatedAt
      })));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(abnormalRows.length ? abnormalRows : [{ 안내: "이상 내역 없음" }]), "07_이상내역");
    XLSX.writeFile(wb, `${year}년_${String(m).padStart(2, "0")}월_PA센터_위생점검.xlsx`);
    flash("월간 점검자료를 생성했습니다.");
  }

  function CalendarStatus({ date, building, kind, label }: { date: string; building: Building; kind: InspectionKind; label: string }) {
    const record = kind === "시설청소" ? weeklyRecord(date, building) : dailyRecord(date, building, kind as "남WC" | "여WC");
    const status = recordStatus(record);
    return <span className={`calendar-task ${status}`}><b>{label}</b></span>;
  }

  function StatusText({ record }: { record?: InspectionRecord }) {
    const status = recordStatus(record);
    return <span className={`status-text ${status}`}>{statusLabel(status)}</span>;
  }

  function FacilityWeekCell({ week, building }: { week: string; building: Building }) {
    const record = weeklyRecord(week, building);
    const status = recordStatus(record);
    const future = isFuture(week);
    const excluded = isFacilityWeekExcluded(week);
    const visualStatus = record ? status : future ? "future" : "missing";
    const label = record ? (status === "normal" ? "점검 완료" : "보완 필요") : excluded ? "점검 제외" : future ? "점검 예정" : "미점검";
    const dateText = record ? formatShort(record.date) : excluded ? "휴무 주간" : future ? "점검 예정" : "미점검";
    const icon = record ? (status === "normal" ? "✓" : "!") : excluded ? "×" : "–";
    const disabled = (future || excluded) && !record;

    return <button
      className={`facility-week-cell ${visualStatus}`}
      disabled={disabled}
      onClick={() => {
        setSelectedDate(week);
        if (record) { setTargetRecordId(record.id); setScreen("detail"); }
        else openTarget(building, "시설청소", week);
      }}
    >
      <span className="facility-status-icon" aria-hidden="true">{icon}</span>
      <strong>{label}</strong>
      <small>{dateText}</small>
    </button>;
  }

  function InspectorNames({ record }: { record?: InspectionRecord }) {
    if (!record) return null;
    return <small>{record.inspectors[0]} · {record.inspectors[1]}</small>;
  }

  return <main className="app-shell">
    {dataLoading && <div className="shared-data-banner loading">공용 점검 데이터를 불러오는 중입니다.</div>}
    {dataSaving && <div className="shared-data-banner saving">공용 데이터에 저장 중입니다.</div>}
    {dataError && <div className="shared-data-banner error">{dataError}</div>}
    <header className="app-header">
      <div className="header-inner">
        <button className="brand-button" onClick={() => setScreen("calendar")} aria-label="점검현황으로 이동">
          <img src="/pulmuone-academy-ci-white.png" alt="풀무원아카데미"/>
        </button>
        <div className="header-actions">
          <nav className="header-nav" aria-label="주 메뉴">
            <button className={screen !== "admin" && screen !== "adminGate" ? "active" : ""} onClick={() => setScreen("calendar")}>점검현황</button>
            <button className={screen === "admin" || screen === "adminGate" ? "active" : ""} onClick={openAdmin}>관리자 접속</button>
          </nav>
        </div>
      </div>
    </header>

    {screen === "calendar" && <section className="page page-calendar">
      <div className="page-heading">
        <span className="eyebrow">PA센터 가치실천팀</span>
        <h1>위생 점검 현황</h1>
        <p>점검일을 확인하고 담당 공간의 위생 상태를 꼼꼼히 점검해 주세요.</p>
      </div>

      <article className="card calendar-card">
        <div className="calendar-head">
          <button onClick={() => setMonth(new Date(month.getFullYear(), month.getMonth() - 1, 1))}>이전</button>
          <strong>{formatMonth(month)}</strong>
          <button onClick={() => setMonth(new Date(month.getFullYear(), month.getMonth() + 1, 1))}>다음</button>
        </div>
        <div className="weekdays">{["일", "월", "화", "수", "목", "금", "토"].map((day) => <span key={day}>{day}</span>)}</div>
        <div className="calendar-grid">
          {calendarDays.map((day, index) => {
            if (!day) return <div className="calendar-empty" key={`empty-${index}`} />;
            const key = toKey(new Date(month.getFullYear(), month.getMonth(), day));
            const customReason = excludedReason(key);
            const closed = isInspectionClosed(key); const future = isFuture(key); const disabled = closed || future;
            return <button
              key={key}
              className={`${selectedDate === key ? "selected" : ""} ${closed ? "closed" : ""} ${future ? "future" : ""}`}
              disabled={disabled}
              onClick={() => setSelectedDate(key)}
            >
              <div className="date-line"><strong>{day}</strong>{(holidayName(key) || customReason) && <em>{holidayName(key) || customReason}</em>}</div>
              {!disabled && <div className="day-status-grid">
                <div className="day-building">
                  <b>본동</b>
                  <CalendarStatus date={key} building="본동" kind="시설청소" label="시설" />
                  <CalendarStatus date={key} building="본동" kind="남WC" label="남WC" />
                  <CalendarStatus date={key} building="본동" kind="여WC" label="여WC" />
                </div>
                <div className="day-building">
                  <b>부속동</b>
                  <CalendarStatus date={key} building="부속동" kind="시설청소" label="시설" />
                  <CalendarStatus date={key} building="부속동" kind="남WC" label="남WC" />
                  <CalendarStatus date={key} building="부속동" kind="여WC" label="여WC" />
                </div>
              </div>}
            </button>;
          })}
        </div>
        <div className="legend">
          <span><i className="normal"/>완료·이상없음</span>
          <span><i className="issue"/>완료·이상있음</span>
          <span><i className="missing"/>미점검</span>
        </div>
      </article>

      <article className="card status-card equal-card">
        <div className="section-head section-head-center">
          <div><span>{weekLabel(selectedDate)} · 주 1회 점검</span><h2>시설 위생 점검 현황</h2></div>
          <button className="text-link" onClick={openFacilityMonthly}>월간 현황</button>
        </div>

        {(["본동", "부속동"] as Building[]).map((building) => {
          const record = weeklyRecord(selectedDate, building);
          return <div className="status-box single-status-box" key={building}><button className="status-row" onClick={() => openTarget(building, "시설청소")}>
            <span className="row-main"><strong>{building} 시설</strong><InspectorNames record={record}/></span>
            <StatusText record={record}/>
          </button></div>;
        })}
      </article>

      <article className="card status-card equal-card">
        <div className="section-head section-head-center"><div><span>{formatDay(selectedDate)}</span><h2>화장실 위생 점검 현황</h2></div><button className="text-link" onClick={() => openRestroomMonthly("본동", "남WC")}>월간 현황</button></div>
        {([
          ["본동", "남WC"],
          ["본동", "여WC"],
          ["부속동", "남WC"],
          ["부속동", "여WC"],
        ] as const).map(([building, kind]) => {
          const record = dailyRecord(selectedDate, building, kind);
          return <div className="status-box single-status-box" key={`${building}-${kind}`}><button className="status-row" onClick={() => openTarget(building, kind)}>
            <span className="row-main"><strong>{building} {restroomLabel(kind)}</strong><InspectorNames record={record}/></span>
            <StatusText record={record}/>
          </button></div>;
        })}
      </article>
    </section>}

    {screen === "inspectors" && <section className="page narrow-page">
      <button className="back-button" onClick={() => setScreen("calendar")}>이전</button>
      <div className="page-heading compact-heading"><span className="eyebrow">INSPECTOR</span><h1>점검자 확인</h1><p>{targetBuilding} {kindLabel(targetKind)} · {targetKind === "시설청소" ? weekLabel(selectedDate) : formatDay(selectedDate)}</p></div>
      <article className="card inspector-card">
        {recentInspectors[0] && recentInspectors[1] && (
          <div className="recent-box"><span>최근 점검자</span><strong>{recentInspectors[0]} · {recentInspectors[1]}</strong></div>
        )}
        <label>
  점검자 1
  <select
    value={inspector1}
    onChange={(e) => setInspector1(e.target.value)}
  >
    <option value="">점검자를 선택하세요</option>
    {activeStaff.map((person) => (
      <option key={person.id} value={person.name}>
        {person.name}
      </option>
    ))}
  </select>
</label>

<label>
  점검자 2
  <select
    value={inspector2}
    onChange={(e) => setInspector2(e.target.value)}
  >
    <option value="">점검자를 선택하세요</option>
    {activeStaff.map((person) => (
      <option key={person.id} value={person.name}>
        {person.name}
      </option>
    ))}
  </select>
</label>
        <button className="primary-button" onClick={startChecklist}>이대로 점검</button>
      </article>
    </section>}

    {screen === "checklist" && <section className="page checklist-page">
      <div className="checklist-top">
        <button className="back-button" onClick={() => targetRecordId ? setScreen("detail") : setScreen("calendar")}>이전</button>
        <div><span>{targetBuilding}</span><h1>{kindLabel(targetKind)} 점검표</h1><p>{targetKind === "시설청소" ? weekLabel(selectedDate) : formatDay(selectedDate)} · {(sessionInspectors ?? [inspector1, inspector2]).join(" · ")}</p></div>
      </div>
      <article className="card checklist-inspector-card">
        <div className="section-head"><div><span>INSPECTORS</span><h2>점검자</h2></div></div>
        <div className="inspector-grid">
          <label>점검자 1<select value={(sessionInspectors ?? [inspector1, inspector2])[0]} onChange={(e) => { const next: [string, string] = [e.target.value, (sessionInspectors ?? [inspector1, inspector2])[1]]; setInspector1(next[0]); setSessionInspectors(next); }}><option value="">점검자를 선택하세요</option>{activeStaff.map((person) => <option key={person.id} value={person.name}>{person.name}</option>)}</select></label>
          <label>점검자 2<select value={(sessionInspectors ?? [inspector1, inspector2])[1]} onChange={(e) => { const next: [string, string] = [(sessionInspectors ?? [inspector1, inspector2])[0], e.target.value]; setInspector2(next[1]); setSessionInspectors(next); }}><option value="">점검자를 선택하세요</option>{activeStaff.map((person) => <option key={person.id} value={person.name}>{person.name}</option>)}</select></label>
        </div>
      </article>
      <article className="card progress-card">
        <div><span>점검 진행</span><strong>{activeChecklist.filter((item) => answers[item.id]).length} / {activeChecklist.length}</strong></div>
        <div className="progress"><i style={{ width: `${activeChecklist.length ? activeChecklist.filter((item) => answers[item.id]).length / activeChecklist.length * 100 : 0}%` }} /></div>
      </article>
      <button className="primary-button all-normal" onClick={markAllNormal}>전체 정상으로 입력</button>
      <div className="checklist">
        {activeChecklist.map((item, index) => <article className={`card check-item ${answers[item.id] === "이상" ? "abnormal" : ""}`} id={`item-${item.id}`} key={item.id}>
          <div className="item-title"><span>{String(index + 1).padStart(2, "0")}</span><strong>{item.name}</strong></div>
          <div className="result-buttons">
            <button className={answers[item.id] === "정상" ? "selected normal" : "normal"} onClick={() => setAnswers((prev) => ({ ...prev, [item.id]: "정상" }))}>정상</button>
            <button className={answers[item.id] === "이상" ? "selected issue" : "issue"} onClick={() => setAnswers((prev) => ({ ...prev, [item.id]: "이상" }))}>이상</button>
          </div>
          {answers[item.id] === "이상" && <textarea value={notes[item.id] ?? ""} onChange={(e) => setNotes((prev) => ({ ...prev, [item.id]: e.target.value }))} placeholder="이상 내용을 입력해 주세요." />}
        </article>)}
      </div>
      <button className="primary-button finish-button" onClick={saveChecklist}>{targetRecordId ? "수정 내용 저장" : "점검 완료"}</button>
    </section>}

    {screen === "detail" && currentRecord && <section className="page detail-page">
      <button className="back-button" onClick={() => setScreen("calendar")}>이전</button>
      <div className="page-heading compact-heading"><span className="eyebrow">INSPECTION RECORD</span><h1>{currentRecord.building} {kindLabel(currentRecord.kind)} 점검표</h1><p>{formatDay(currentRecord.date)}</p></div>
      <article className="card record-summary">
        <div><span>점검자</span><strong>{currentRecord.inspectors.join(" · ")}</strong></div>
        <div><span>점검 결과</span><StatusText record={currentRecord}/></div>
        <div><span>최종 저장</span><strong>{currentRecord.updatedAt}</strong></div>
      </article>
      <div className="detail-list">
        {currentRecord.items.map((item, index) => <article className={`card detail-item ${item.result === "이상" ? "abnormal" : ""}`} key={`${item.itemId}-${index}`}>
          <span>{String(index + 1).padStart(2, "0")}</span>
          <div><strong>{item.name}</strong>{item.note && <p>{item.note}</p>}</div>
          <b className={item.result === "이상" ? "issue-text" : "normal-text"}>{item.result}</b>
        </article>)}
      </div>
      <button className="primary-button finish-button" onClick={beginEdit}>수정하기</button>
    </section>}

    {screen === "restroomMonthly" && <section className="page monthly-page">
      <button className="back-button" onClick={() => setScreen("calendar")}>이전</button>
      <div className="page-heading compact-heading"><span className="eyebrow">MONTHLY STATUS</span><h1>화장실 월간 위생 점검 현황</h1><p>장소를 선택해 점검항목별 월간 O/X 현황을 확인할 수 있습니다.</p></div>
      <article className="card monthly-card">
        <div className="segment-tabs restroom-monthly-tabs">
          {(["본동", "부속동"] as Building[]).map((building) => <button key={building} className={targetBuilding === building ? "active" : ""} onClick={() => setTargetBuilding(building)}>{building}</button>)}
        </div>
        <div className="segment-tabs restroom-monthly-tabs gender-tabs">
          {(["남WC", "여WC"] as const).map((kind) => <button key={kind} className={targetKind === kind ? "active" : ""} onClick={() => setTargetKind(kind)}>{kind === "남WC" ? "남자 화장실" : "여자 화장실"}</button>)}
        </div>
        <div className="calendar-head"><button onClick={() => setMonth(new Date(month.getFullYear(), month.getMonth() - 1, 1))}>이전</button><strong>{formatMonth(month)} · {targetBuilding} {kindLabel(targetKind)}</strong><button onClick={() => setMonth(new Date(month.getFullYear(), month.getMonth() + 1, 1))}>다음</button></div>
        <div className="monthly-table-wrap">
          <table className="monthly-table"><thead><tr><th>점검항목</th>{Array.from({ length: new Date(month.getFullYear(), month.getMonth() + 1, 0).getDate() }, (_, index) => <th key={index + 1}>{index + 1}</th>)}</tr></thead>
          <tbody>{managedItems.restroom.map((managed) => <tr key={managed.id}><th>{managed.name}</th>{Array.from({ length: new Date(month.getFullYear(), month.getMonth() + 1, 0).getDate() }, (_, index) => {
            const key = toKey(new Date(month.getFullYear(), month.getMonth(), index + 1));
            if (isInspectionClosed(key)) return <td className="closed-cell" key={key}>·</td>;
            const record = dailyRecord(key, targetBuilding, targetKind as "남WC" | "여WC");
            const item = record?.items.find((target) => target.name === managed.name);
            return <td key={key} className={item?.result === "이상" ? "x-cell" : item?.result === "정상" ? "o-cell" : "missing-cell"} onClick={() => { if (record) { setSelectedDate(key); setTargetRecordId(record.id); setScreen("detail"); } }}>{item ? (item.result === "정상" ? "O" : "X") : "-"}</td>;
          })}</tr>)}</tbody></table>
        </div>
      </article>
    </section>}

    {screen === "facilityMonthly" && <section className="page monthly-page facility-monthly-page">
      <button className="back-button" onClick={() => setScreen("calendar")}>이전</button>
      <div className="monthly-title-row">
        <div className="page-heading compact-heading"><span className="eyebrow">MONTHLY STATUS</span><h1>시설 월간 위생 점검 현황</h1><p>화장실 항목은 제외하고, 본동·부속동 시설 점검 결과만 주차별로 확인합니다.</p></div>
      </div>
      <article className="card monthly-card facility-monthly-card">
        <div className="calendar-head"><button onClick={() => setMonth(new Date(month.getFullYear(), month.getMonth() - 1, 1))}>이전</button><strong>{formatMonth(month)}</strong><button onClick={() => setMonth(new Date(month.getFullYear(), month.getMonth() + 1, 1))}>다음</button></div>
        <div className="facility-monthly-scroll">
          <div className="facility-monthly-grid" style={{ gridTemplateColumns: `minmax(150px, .9fr) repeat(${weeksForMonth(month).length}, minmax(140px, 1fr))` }}>
            <div className="facility-grid-head facility-grid-corner"><strong>구분</strong></div>
            {weeksForMonth(month).map((week, index, arr) => <div className={`facility-grid-head ${index === arr.length - 1 ? "col-end" : ""}`} key={week}>
              <strong>{index + 1}주차</strong>
              <span>{weekLabel(week)}</span>
            </div>)}
            {(["본동", "부속동"] as Building[]).map((building) => <>
              <div className={`facility-building-cell ${building === "부속동" ? "row-end" : ""}`} key={`${building}-label`}>
                <strong>{building}</strong>
              </div>
              {weeksForMonth(month).map((week, weekIndex, weekArr) => <div className={`facility-grid-cell ${weekIndex === weekArr.length - 1 ? "col-end" : ""} ${building === "부속동" ? "row-end" : ""}`} key={`${building}-${week}`}>
                <FacilityWeekCell week={week} building={building} />
              </div>)}
            </>)}
          </div>
        </div>
        <div className="facility-status-legend">
          <span className="normal"><i>✓</i>점검 완료</span>
          <span className="issue"><i>!</i>보완 필요</span>
          <span className="missing"><i>–</i>미점검</span>
          <span className="future"><i>–</i>점검 예정</span>
        </div>
        <p className="facility-monthly-note">주차 기준은 월요일 ~ 금요일이며, 각 칸을 선택하면 해당 주차의 시설 점검 등록 또는 상세 화면으로 이동합니다.</p>
      </article>
    </section>}

    {screen === "adminGate" && <section className="page admin-gate-page">
      <article className="card admin-access-card">
        <div className="page-heading admin-access-heading"><span className="eyebrow">ADMIN ACCESS</span><h1>관리자 접속</h1><p>비밀번호 인증 후 점검자 관리, 점검항목 관리, 월간 자료출력 기능을 사용할 수 있습니다.</p></div>
        <div className="card admin-access-form-card">
          <label className="admin-access-label">비밀번호
            <input type="password" value={adminInput} onChange={(e) => setAdminInput(e.target.value)} placeholder="관리자 비밀번호를 입력하세요" onKeyDown={(e) => { if (e.key === "Enter") submitAdminAccess(); }} />
          </label>
          <div className="admin-access-actions">
            <button className="secondary-button admin-access-button" onClick={() => setScreen("calendar")}>점검현황으로 돌아가기</button>
            <button className="primary-button admin-access-button" onClick={submitAdminAccess}>접속하기</button>
          </div>
        </div>
      </article>
    </section>}

    {screen === "admin" && <section className="page admin-page">
      <div className="admin-page-top">
        <div className="page-heading"><span className="eyebrow">ADMIN</span><h1>관리자</h1><p>점검자와 점검항목을 관리하고 월간 자료를 한 번에 출력합니다.</p></div>
        <button className="secondary-button admin-logout-button" onClick={logoutAdmin}>로그아웃</button>
      </div>
      <div className="admin-tabs"><button className={adminTab === "staff" ? "active" : ""} onClick={() => setAdminTab("staff")}>점검자 관리</button><button className={adminTab === "items" ? "active" : ""} onClick={() => setAdminTab("items")}>점검항목 관리</button><button className={adminTab === "closedDates" ? "active" : ""} onClick={() => setAdminTab("closedDates")}>점검 제외일</button><button className={adminTab === "export" ? "active" : ""} onClick={() => setAdminTab("export")}>자료출력</button><button className={adminTab === "security" ? "active" : ""} onClick={() => setAdminTab("security")}>관리자 보안</button></div>

      {adminTab === "staff" && <article className="card admin-card"><div className="section-head"><div><span>INSPECTORS</span><h2>점검자 관리</h2></div><button className="secondary-button" onClick={startAddStaff} disabled={staffAdding}>점검자 추가</button></div><div className="manage-list">{staff.map((person) => staffEditId === person.id ? <div className="manage-row inline-editor-row" key={person.id}><div className="inline-editor-main"><input autoFocus value={staffDraft} onChange={(e) => setStaffDraft(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter") saveStaffEditor(); if (e.key === "Escape") cancelStaffEditor(); }} aria-label="점검자 이름 수정" /></div><div className="inline-editor-actions"><button className="save-inline-button" onClick={saveStaffEditor}>저장</button><button onClick={cancelStaffEditor}>취소</button></div></div> : <div className="manage-row" key={person.id}><div><strong>{person.name}</strong><span>{person.active ? "사용 중" : "사용 중지"}</span></div><div><button onClick={() => startEditStaff(person)}>수정</button><button onClick={() => toggleStaff(person)}>{person.active ? "사용 중지" : "다시 사용"}</button><button className="danger-button" onClick={() => deleteStaff(person)}>삭제</button></div></div>)}{staffAdding && <div className="manage-row inline-editor-row add-inline-row"><div className="inline-editor-main"><input autoFocus value={staffDraft} onChange={(e) => setStaffDraft(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter") saveStaffEditor(); if (e.key === "Escape") cancelStaffEditor(); }} placeholder="새 점검자 이름을 입력하세요" aria-label="새 점검자 이름" /></div><div className="inline-editor-actions"><button className="save-inline-button" onClick={saveStaffEditor}>저장</button><button onClick={cancelStaffEditor}>취소</button></div></div>}</div></article>}

      {adminTab === "items" && <article className="card admin-card"><div className="section-head"><div><span>CHECKLIST</span><h2>점검항목 관리</h2></div><button className="secondary-button" onClick={startAddChecklistItem} disabled={itemAdding}>항목 추가</button></div><div className="segment-tabs"><button className={itemTab === "main" ? "active" : ""} onClick={() => { cancelItemEditor(); setItemTab("main"); }}>본동 시설</button><button className={itemTab === "annex" ? "active" : ""} onClick={() => { cancelItemEditor(); setItemTab("annex"); }}>부속동 시설</button><button className={itemTab === "restroom" ? "active" : ""} onClick={() => { cancelItemEditor(); setItemTab("restroom"); }}>화장실 공통</button></div><p className="admin-note">항목 변경은 새 점검부터 적용되며, 저장된 과거 점검표는 당시 항목 그대로 유지됩니다.</p><div className="manage-list item-manage-list">{managedItems[itemTab].map((item, index) => itemEditId === item.id ? <div className="manage-row inline-editor-row item-inline-editor" key={item.id}><div className="item-name inline-item-name"><b>{index + 1}</b><input autoFocus value={itemDraft} onChange={(e) => setItemDraft(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter") saveChecklistItemEditor(); if (e.key === "Escape") cancelItemEditor(); }} aria-label="점검항목 수정" /></div><div className="inline-editor-actions"><button className="save-inline-button" onClick={saveChecklistItemEditor}>저장</button><button onClick={cancelItemEditor}>취소</button></div></div> : <div className="manage-row" key={item.id}><div className="item-name"><b>{index + 1}</b><span><strong>{item.name}</strong><small>{item.active ? "사용 중" : "사용 중지"}</small></span></div><div><button disabled={index === 0} onClick={() => moveChecklistItem(index, -1)}>위</button><button disabled={index === managedItems[itemTab].length - 1} onClick={() => moveChecklistItem(index, 1)}>아래</button><button onClick={() => startEditChecklistItem(item)}>수정</button><button onClick={() => toggleChecklistItem(item)}>{item.active ? "사용 중지" : "다시 사용"}</button><button className="danger-button" onClick={() => deleteChecklistItem(item)}>삭제</button></div></div>)}{itemAdding && <div className="manage-row inline-editor-row item-inline-editor add-inline-row"><div className="item-name inline-item-name"><b>{managedItems[itemTab].length + 1}</b><input autoFocus value={itemDraft} onChange={(e) => setItemDraft(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter") saveChecklistItemEditor(); if (e.key === "Escape") cancelItemEditor(); }} placeholder="새 점검항목을 입력하세요" aria-label="새 점검항목" /></div><div className="inline-editor-actions"><button className="save-inline-button" onClick={saveChecklistItemEditor}>저장</button><button onClick={cancelItemEditor}>취소</button></div></div>}</div></article>}


      {adminTab === "closedDates" && <article className="card admin-card excluded-date-card">
        <div className="section-head"><div><span>EXCLUDED DATES</span><h2>점검 제외일 관리</h2></div></div>
        <p className="admin-note">단체휴무 등 수기 점검을 하지 않는 날짜를 등록합니다. 등록일은 캘린더에 사유가 표시되고 점검 대상 및 미점검 집계에서 제외됩니다.</p>
        <div className="excluded-date-editor">
          <label>날짜<input type="date" value={excludedDateDraft} onChange={(e) => setExcludedDateDraft(e.target.value)} /></label>
          <label>제외 사유<input value={excludedReasonDraft} onChange={(e) => setExcludedReasonDraft(e.target.value)} placeholder="예: 전사 단체휴무" onKeyDown={(e) => { if (e.key === "Enter") saveExcludedDate(); }} /></label>
          <div className="excluded-date-actions"><button className="primary-button" onClick={saveExcludedDate}>{excludedEditingDate ? "수정 저장" : "제외일 추가"}</button>{excludedEditingDate && <button className="secondary-button" onClick={resetExcludedDateEditor}>취소</button>}</div>
        </div>
        <div className="manage-list excluded-date-list">{excludedDates.length ? excludedDates.map((item) => <div className="manage-row" key={item.date}><div><strong>{item.date}</strong><span>{item.reason}</span></div><div><button onClick={() => startEditExcludedDate(item)}>수정</button><button className="danger-button" onClick={() => deleteExcludedDate(item.date)}>삭제</button></div></div>) : <div className="empty-admin-list">등록된 점검 제외일이 없습니다.</div>}</div>
      </article>}

      {adminTab === "export" && <article className="card admin-card export-card"><div className="section-head"><div><span>MONTHLY EXPORT</span><h2>월간 점검자료 일괄 출력</h2></div></div><p>선택한 월의 시설 및 화장실 위생 점검자료를 하나의 Excel 파일로 내려받습니다.</p><label>출력 월<input type="month" value={exportMonth} onChange={(e) => setExportMonth(e.target.value)} /></label><div className="export-sheet-list"><span>01_본동시설</span><span>02_부속동시설</span><span>03_본동_남자화장실</span><span>04_본동_여자화장실</span><span>05_부속동_남자화장실</span><span>06_부속동_여자화장실</span><span>07_이상내역</span></div><button className="primary-button" onClick={exportMonthlyWorkbook}>월간 전체 Excel 다운로드</button></article>}

      {adminTab === "security" && <article className="card admin-card security-card"><div className="section-head"><div><span>ADMIN SECURITY</span><h2>관리자 비밀번호 설정</h2></div></div><p className="admin-note">관리자 화면 진입 시 사용하는 비밀번호를 변경할 수 있습니다. 현재 브라우저에 저장되며 같은 기기에서 유지됩니다.</p><div className="security-grid"><label>현재 비밀번호<input type="password" value={securityCurrent} onChange={(e) => setSecurityCurrent(e.target.value)} placeholder="현재 비밀번호" /></label><label>새 비밀번호<input type="password" value={securityNext} onChange={(e) => setSecurityNext(e.target.value)} placeholder="새 비밀번호" /></label><label>새 비밀번호 확인<input type="password" value={securityConfirm} onChange={(e) => setSecurityConfirm(e.target.value)} placeholder="새 비밀번호 확인" /></label></div><div className="security-actions"><button className="primary-button" onClick={updateAdminPassword}>비밀번호 저장</button></div></article>}
    </section>}

  </main>;
}
