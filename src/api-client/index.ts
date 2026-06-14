/**
 * Local mock for @workspace/api-client-react
 * All data is stored in localStorage so it persists across page refreshes.
 * Replace this file (or its fetch calls) with real API calls when you have a backend.
 */

import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";

// ─── Types ────────────────────────────────────────────────────────────────────

export interface CoreData {
  ratio?: string;
  burden_va?: string;
  accuracy_class?: string;
  isf?: string;
  min_knee_pt_volt?: string;
  max_rct_75c?: string;
  max_exc_vk2?: string;
  bare_core_dim?: string;
  core_material?: string;
  core_weight_kg?: string;
  sec_total_turns?: string;
  sec_ter_marking?: string;
  sec_cond_s1s2?: string;
  sec_turns_s1s2?: string;
  sec_cond_s2s3?: string;
  sec_turns_s2s3?: string;
  sec_cond_s3s4?: string;
  sec_turns_s3s4?: string;
  sec_cond_s4s5?: string;
  sec_turns_s4s5?: string;
  sec_copper_wt?: string;
  finished_core_dim?: string;
  sec_connection?: string;
  wire_length?: string;
  wire_colour?: string;
  [key: string]: string | undefined;
}

export interface ItemInput {
  item_no: string;
  ct_type?: string;
  cust_part_code?: string;
  ratio?: string;
  rated_voltage?: string;
  stc?: string;
  insulation_level?: string;
  frequency?: string;
  ref_std?: string;
  core1?: CoreData;
  core2?: CoreData;
  core3?: CoreData;
  ct_final_dim?: string;
  ga_drg?: string;
  ins_class?: string;
  ref_ti?: string;
  pri_turns?: string;
  pri_copper?: string;
  former?: string;
  pri_length?: string;
  pri_weight?: string;
  sec_terminal?: string;
  total_weight?: string;
}

export interface TiRecordInput {
  item_no?: string;
  ti_no?: string;
  ti_date?: string;
  wo_no?: string;
  customer_name?: string;
  qty?: string;
  ct_type?: string;
  cust_part_code?: string;
  ratio?: string;
  rated_voltage?: string;
  stc?: string;
  insulation_level?: string;
  frequency?: string;
  ref_std?: string;
  core1?: CoreData;
  core2?: CoreData;
  core3?: CoreData;
  ct_final_dim?: string;
  ga_drg?: string;
  ins_class?: string;
  ref_ti?: string;
  pri_turns?: string;
  pri_copper?: string;
  former?: string;
  pri_length?: string;
  pri_weight?: string;
  sec_terminal?: string;
  total_weight?: string;
  created_by?: string;
  checked_by?: string;
  approved_by?: string;
  remarks?: string;
  rev_no?: string;
  note?: string;
  [key: string]: unknown;
}

interface TiRecord extends TiRecordInput {
  id: string;
  ti_no: string;
}

interface Item extends ItemInput {
  id: string;
}

// ─── localStorage helpers ─────────────────────────────────────────────────────

function getItems(): Item[] {
  try {
    return JSON.parse(localStorage.getItem("ct_items") || "[]");
  } catch {
    return [];
  }
}

function setItems(items: Item[]) {
  localStorage.setItem("ct_items", JSON.stringify(items));
}

function getTiRecords(): TiRecord[] {
  try {
    return JSON.parse(localStorage.getItem("ct_ti_records") || "[]");
  } catch {
    return [];
  }
}

function setTiRecords(records: TiRecord[]) {
  localStorage.setItem("ct_ti_records", JSON.stringify(records));
}

function getTiCounter(): number {
  return parseInt(localStorage.getItem("ct_ti_counter") || "0", 10);
}

function incrementTiCounter(): number {
  const next = getTiCounter() + 1;
  localStorage.setItem("ct_ti_counter", String(next));
  return next;
}

function generateTiNo(): string {
  const now = new Date();
  const m = now.getMonth() + 1;
  const y = now.getFullYear() % 100;
  const fyS = m >= 4 ? y : y - 1;
  const fyE = fyS + 1;
  const fy = `${String(fyS).padStart(2, "0")}-${String(fyE).padStart(2, "0")}`;
  const seq = incrementTiCounter();
  return `TI/${fy}/${String(seq).padStart(4, "0")}`;
}

// ─── Query key factories ──────────────────────────────────────────────────────

export function getGetItemQueryKey(itemNo: string) {
  return ["item", itemNo];
}

export function getGetTiRecordQueryKey(tiNo: string) {
  return ["ti-record", tiNo];
}

// ─── Hooks ────────────────────────────────────────────────────────────────────

export function useGetItem(
  itemNo: string,
  options?: { query?: { enabled?: boolean } }
) {
  return useQuery({
    queryKey: getGetItemQueryKey(itemNo),
    queryFn: () => {
      const items = getItems();
      const found = items.find((i) => i.item_no === itemNo);
      if (!found) throw new Error("Item not found");
      return found;
    },
    enabled: options?.query?.enabled !== false && !!itemNo,
    retry: false,
  });
}

export function useGetTiRecord(
  tiNo: string,
  options?: { query?: { enabled?: boolean } }
) {
  return useQuery({
    queryKey: getGetTiRecordQueryKey(tiNo),
    queryFn: () => {
      const records = getTiRecords();
      const found = records.find((r) => r.ti_no === tiNo);
      if (!found) throw new Error("TI record not found");
      return found;
    },
    enabled: options?.query?.enabled !== false && !!tiNo,
    retry: false,
  });
}

export function useGetAdjacentTiRecords(
  tiNo: string,
  options?: { query?: { enabled?: boolean } }
) {
  return useQuery({
    queryKey: ["ti-adjacent", tiNo],
    queryFn: () => {
      const records = getTiRecords();
      const idx = records.findIndex((r) => r.ti_no === tiNo);
      return {
        prev: idx > 0 ? records[idx - 1].ti_no : null,
        next: idx < records.length - 1 ? records[idx + 1].ti_no : null,
      };
    },
    enabled: options?.query?.enabled !== false && !!tiNo,
    retry: false,
  });
}

export function useListTiRecords(
  filters: {
    tiNo?: string;
    itemNo?: string;
    customer?: string;
    woNo?: string;
    ctType?: string;
  },
  options?: { query?: { enabled?: boolean } }
) {
  return useQuery({
    queryKey: ["ti-records", filters],
    queryFn: () => {
      let records = getTiRecords();
      if (filters.tiNo) {
        records = records.filter((r) =>
          r.ti_no?.toLowerCase().includes(filters.tiNo!.toLowerCase())
        );
      }
      if (filters.itemNo) {
        records = records.filter((r) =>
          r.item_no?.toLowerCase().includes(filters.itemNo!.toLowerCase())
        );
      }
      if (filters.customer) {
        records = records.filter((r) =>
          r.customer_name?.toLowerCase().includes(filters.customer!.toLowerCase())
        );
      }
      if (filters.woNo) {
        records = records.filter((r) =>
          r.wo_no?.toLowerCase().includes(filters.woNo!.toLowerCase())
        );
      }
      if (filters.ctType) {
        records = records.filter((r) =>
          r.ct_type?.toLowerCase().includes(filters.ctType!.toLowerCase())
        );
      }
      return { records };
    },
    enabled: options?.query?.enabled !== false,
    retry: false,
  });
}

export function useGenerateTiNumber() {
  return useMutation({
    mutationFn: async (_args: Record<string, never>) => {
      return { ti_no: generateTiNo() };
    },
  });
}

export function useCreateTiRecord() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({ data }: { data: TiRecordInput }) => {
      const records = getTiRecords();
      const tiNo = data.ti_no || generateTiNo();
      const newRecord: TiRecord = {
        ...data,
        id: crypto.randomUUID(),
        ti_no: tiNo,
        ti_date:
          data.ti_date ||
          new Date().toLocaleDateString("en-GB", {
            day: "2-digit",
            month: "2-digit",
            year: "numeric",
          }),
      };
      records.push(newRecord);
      setTiRecords(records);
      return newRecord;
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["ti-records"] });
    },
  });
}

export function useUpdateTiRecord() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({
      tiNo,
      data,
    }: {
      tiNo: string | null;
      data: TiRecordInput;
    }) => {
      const records = getTiRecords();
      const idx = records.findIndex((r) => r.ti_no === tiNo);
      if (idx === -1) throw new Error("TI record not found");
      records[idx] = { ...records[idx], ...data, ti_no: tiNo! };
      setTiRecords(records);
      return records[idx];
    },
    onSuccess: (_data, variables) => {
      queryClient.invalidateQueries({ queryKey: ["ti-records"] });
      queryClient.invalidateQueries({
        queryKey: getGetTiRecordQueryKey(variables.tiNo || ""),
      });
    },
  });
}

export function useCreateItem() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({ data }: { data: ItemInput }) => {
      const items = getItems();
      const newItem: Item = { ...data, id: crypto.randomUUID() };
      items.push(newItem);
      setItems(items);
      return newItem;
    },
    onSuccess: (_data, variables) => {
      queryClient.invalidateQueries({
        queryKey: getGetItemQueryKey(variables.data.item_no),
      });
    },
  });
}
