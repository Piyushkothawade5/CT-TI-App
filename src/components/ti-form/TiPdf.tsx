import React from "react";
import { Document, Page, View, Text, StyleSheet, Font } from "@react-pdf/renderer";
import type { TiRecordInput, CoreData } from "@/api-client";

// ─────────────────────────────────────────────────────────────────────────────
// A4 CONSTANTS  (points: 1pt = 1/72 inch)
// ─────────────────────────────────────────────────────────────────────────────
const A4_H = 841.89;   // pt
const A4_W = 595.28;   // pt

// Fixed structural heights (pt) — measured empirically from the layout below
const FIXED_HEIGHTS = {
  pageTopPad:    18,
  pageBottomPad: 18,
  header:        46,   // headerRow incl. border
  infoTable:     5 * 20,  // 5 rows × minHeight 20 ≈ 100
  electricRow:   24,
  coreTableHead: 18,
  noteTable:     42,   // noteTable minHeight 36 + marginTop 6
  sigRow:        72,   // name 9pt + 18gap + line + label + marginTop 18
  coreMarginTop: 4,
};

const TOTAL_FIXED =
  FIXED_HEIGHTS.pageTopPad +
  FIXED_HEIGHTS.pageBottomPad +
  FIXED_HEIGHTS.header +
  FIXED_HEIGHTS.infoTable +
  FIXED_HEIGHTS.electricRow +
  FIXED_HEIGHTS.coreTableHead +
  FIXED_HEIGHTS.noteTable +
  FIXED_HEIGHTS.sigRow +
  FIXED_HEIGHTS.coreMarginTop;

// Separator rows: 2 separators (sep at index 7) each 6 pt high
// But we count them below in buildRows so we keep it generic.

// ─────────────────────────────────────────────────────────────────────────────
// SCALING ENGINE
// ─────────────────────────────────────────────────────────────────────────────

interface ScaleConfig {
  fontSize: number;       // body font size (pt)
  labelSize: number;      // label/bold size
  headerTitleSize: number;
  headerSubSize: number;
  rowHeight: number;      // minHeight of each core row
  sepHeight: number;      // separator row height
  padH: number;           // cell paddingHorizontal
  padV: number;           // cell paddingVertical
  pagePadH: number;       // page side padding
  pagePadV: number;       // page top/bottom padding
}

/** Estimate total page height for a given scale and row count */
function estimateHeight(scale: ScaleConfig, dataRows: number, sepRows: number): number {
  const coreBodyHeight = dataRows * scale.rowHeight + sepRows * scale.sepHeight;

  const headerH   = scale.headerTitleSize + 4 + scale.headerSubSize + 2 * scale.padV + 2;
  const infoRowH  = scale.fontSize + scale.labelSize + scale.padV * 2 + 3;
  const infoH     = 5 * infoRowH;
  const elecH     = scale.fontSize + scale.labelSize + scale.padV * 2 + 3 + 2;
  const noteH     = 36 + 6; // note minHeight + marginTop
  const sigH      = scale.fontSize * 1.2 + 18 + 4 + scale.fontSize + 18; // name + gap + line + label + marginTop

  return (
    scale.pagePadV * 2 +
    headerH +
    infoH +
    elecH +
    scale.coreMarginTop +
    18 +       // core table header row
    coreBodyHeight +
    noteH +
    sigH
  );
}
(estimateHeight as any).coreMarginTop = 4;  // just to avoid unused warning

/** Binary-search a scale that makes content fit in A4_H */
function computeScale(dataRows: number, sepRows: number): ScaleConfig {
  // Candidate scales from largest to smallest
  const candidates: ScaleConfig[] = [
    // Scale A – roomy (inline info rows free up ~50pt vs stacked)
    { fontSize: 9,   labelSize: 8.5, headerTitleSize: 12, headerSubSize: 10, rowHeight: 14, sepHeight: 6,  padH: 5, padV: 4, pagePadH: 22, pagePadV: 18 },
    // Scale B – comfortable default
    { fontSize: 8.5, labelSize: 8,   headerTitleSize: 11, headerSubSize: 9,  rowHeight: 14, sepHeight: 6,  padH: 4, padV: 3, pagePadH: 22, pagePadV: 16 },
    // Scale C – slightly tighter
    { fontSize: 8,   labelSize: 7.5, headerTitleSize: 10, headerSubSize: 8.5,rowHeight: 13, sepHeight: 5,  padH: 4, padV: 2, pagePadH: 22, pagePadV: 14 },
    // Scale D – compact
    { fontSize: 7.5, labelSize: 7,   headerTitleSize: 9.5,headerSubSize: 8,  rowHeight: 12, sepHeight: 4,  padH: 3, padV: 2, pagePadH: 20, pagePadV: 12 },
    // Scale E – very compact
    { fontSize: 7,   labelSize: 6.5, headerTitleSize: 9,  headerSubSize: 7.5,rowHeight: 11, sepHeight: 4,  padH: 3, padV: 1, pagePadH: 18, pagePadV: 10 },
    // Scale F – ultra-compact (last resort)
    { fontSize: 6.5, labelSize: 6,   headerTitleSize: 8.5,headerSubSize: 7,  rowHeight: 10, sepHeight: 3,  padH: 2, padV: 1, pagePadH: 16, pagePadV: 8  },
  ];

  for (const scale of candidates) {
    const est = estimatePageHeight(scale, dataRows, sepRows);
    if (est <= A4_H) return scale;
  }
  // Fallback – return smallest scale regardless
  return candidates[candidates.length - 1];
}

/** More careful estimation used in the engine */
function estimatePageHeight(scale: ScaleConfig, dataRows: number, sepRows: number): number {
  const lineH = scale.fontSize * 1.2;
  // Info rows are now single-line "LABEL : value" → only one line height needed
  const infoRowH  = Math.max(lineH, scale.labelSize * 1.2) + scale.padV * 2 + 1;
  const headerH   = scale.headerTitleSize * 1.2 + scale.headerSubSize * 1.2 + scale.padV * 2 + 4;
  const elecH     = scale.labelSize * 1.2 + lineH + scale.padV * 2 + 2;
  const coreHeadH = scale.fontSize * 1.2 + scale.padV * 2 + 2;
  const coreBodyH = dataRows * scale.rowHeight + sepRows * scale.sepHeight;
  const noteH     = Math.max(36, scale.labelSize * 1.2 + lineH * 3 + scale.padV * 2) + 6;
  const sigH      = scale.fontSize * 1.1 + 18 + 4 + scale.fontSize + 18 + 4;

  return (
    scale.pagePadV * 2 +
    headerH +
    5 * infoRowH +
    elecH +
    4 +             // coreMarginTop
    coreHeadH +
    coreBodyH +
    noteH +
    sigH
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────────────────────────────────────

const BLACK = "#000000";

function v(val?: string | null): string { return val ?? ""; }

function formatDate(dateStr?: string | null): string {
  if (!dateStr) return "";
  try {
    const d = new Date(dateStr + "T00:00:00");
    const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
    return `${String(d.getDate()).padStart(2,"0")}-${months[d.getMonth()]}-${d.getFullYear()}`;
  } catch { return dateStr; }
}

// ─────────────────────────────────────────────────────────────────────────────
// ROW DEFINITIONS
// ─────────────────────────────────────────────────────────────────────────────

type CoreRow =
  | { type: "core";   label: string; key: keyof CoreData }
  | { type: "single"; label: string; value: string }
  | { type: "separator" };

function buildRows(data: TiRecordInput, maxExcLabel: string): CoreRow[] {
  return [
    { type: "core",   label: "RATIO",                    key: "ratio" },
    { type: "core",   label: "Burden (VA)",               key: "burden_va" },
    { type: "core",   label: "Accuracy Class",            key: "accuracy_class" },
    { type: "core",   label: "ISF",                       key: "isf" },
    { type: "core",   label: "Min. Knee pt. volt.",       key: "min_knee_pt_volt" },
    { type: "core",   label: "Max. Rct @ 75\u00b0c",     key: "max_rct_75c" },
    { type: "core",   label: maxExcLabel,                 key: "max_exc_vk2" },
    { type: "separator" },
    { type: "core",   label: "Core Dimensions",           key: "bare_core_dim" },
    { type: "core",   label: "Core Material",             key: "core_material" },
    { type: "core",   label: "Core weight (Kg)",          key: "core_weight_kg" },
    { type: "core",   label: "Sec. Total Turns",          key: "sec_total_turns" },
    { type: "core",   label: "Sec. Ter. Marking",         key: "sec_ter_marking" },
    { type: "core",   label: "Sec. Conductor (S1-S2)",    key: "sec_cond_s1s2" },
    { type: "core",   label: "Sec. Turns (S1-S2)",        key: "sec_turns_s1s2" },
    { type: "core",   label: "Sec. Conductor (S2-S3)",    key: "sec_cond_s2s3" },
    { type: "core",   label: "Sec. Turns (S2-S3)",        key: "sec_turns_s2s3" },
    { type: "core",   label: "Sec. Conductor (S3-S4)",    key: "sec_cond_s3s4" },
    { type: "core",   label: "Sec. Turns (S3-S4)",        key: "sec_turns_s3s4" },
    { type: "core",   label: "Sec. Conductor (S4-S5)",    key: "sec_cond_s4s5" },
    { type: "core",   label: "Sec. Turns (S4-S5)",        key: "sec_turns_s4s5" },
    { type: "core",   label: "Sec. Copper weight (kg)",   key: "sec_copper_wt" },
    { type: "core",   label: "Finished Core Dim. (mm)",   key: "finished_core_dim" },
    { type: "core",   label: "Sec Connection",            key: "sec_connection" },
    { type: "core",   label: "Wire Length",               key: "wire_length" },
    { type: "core",   label: "Wire Colour",               key: "wire_colour" },
    { type: "single", label: "CT final dim",   value: v(data.ct_final_dim) },
    { type: "single", label: "GA Drg",         value: v(data.ga_drg) },
    { type: "single", label: "INS CLASS",       value: v(data.ins_class) },
    { type: "single", label: "PRI Turns",       value: v(data.pri_turns) },
    { type: "single", label: "PRI Copper",      value: v(data.pri_copper) },
    { type: "single", label: "Former",          value: v(data.former) },
    { type: "single", label: "PRI Length",      value: v(data.pri_length) },
    { type: "single", label: "PRI Weight",      value: v(data.pri_weight) },
    { type: "single", label: "Sec. Terminal",   value: v(data.sec_terminal) },
    { type: "single", label: "Total Weight",    value: v(data.total_weight) },
    { type: "single", label: "Ref TI:",         value: v(data.ref_ti) },
  ];
}

// ─────────────────────────────────────────────────────────────────────────────
// DYNAMIC STYLESHEET FACTORY
// ─────────────────────────────────────────────────────────────────────────────

function makeStyles(sc: ScaleConfig) {
  const B = { fontFamily: "Helvetica-Bold" } as const;

  return StyleSheet.create({
    page: {
      fontFamily: "Helvetica",
      fontSize: sc.fontSize,
      paddingTop: sc.pagePadV,
      paddingBottom: sc.pagePadV,
      paddingLeft: sc.pagePadH,
      paddingRight: sc.pagePadH,
      color: BLACK,
    },

    // ── Header ──
    headerRow: {
      flexDirection: "row",
      borderWidth: 1, borderColor: BLACK, borderStyle: "solid",
      paddingHorizontal: sc.padH + 4,
      paddingVertical: sc.padV + 2,
      marginBottom: 0,
      alignItems: "center",
    },
    headerLeft:  { flex: 1, alignItems: "center" },
    headerTitle: { fontSize: sc.headerTitleSize, ...B, letterSpacing: 0.5 },
    headerSub:   { fontSize: sc.headerSubSize, marginTop: 2 },
    headerRight: { flex: 1, alignItems: "flex-end" },
    headerField: { flexDirection: "row", marginBottom: 3 },
    headerLabel: { ...B, marginRight: 4 },

    // ── Info table ──
    infoTable: {
      borderLeftWidth: 1, borderRightWidth: 1,
      borderBottomWidth: 1, borderTopWidth: 0,
      borderColor: BLACK, borderStyle: "solid",
    },
    infoRow:     { flexDirection: "row", borderBottomWidth: 1, borderColor: BLACK, borderStyle: "solid" },
    infoRowLast: { flexDirection: "row" },
    infoCell: {
      flex: 1,
      paddingHorizontal: sc.padH + 1,
      paddingVertical: sc.padV,
      borderRightWidth: 1, borderColor: BLACK, borderStyle: "solid",
    },
    infoCellLast: {
      flex: 1,
      paddingHorizontal: sc.padH + 1,
      paddingVertical: sc.padV,
    },
    // Inline "LABEL : value" — single row, no stacking
    infoInlineRow: {
      flexDirection: "row",
      alignItems: "center",
    },
    infoLabel:     { ...B, fontSize: sc.labelSize },
    infoColon:     { ...B, fontSize: sc.labelSize, marginHorizontal: 2 },
    infoValue:     { fontSize: sc.fontSize, flex: 1 },

    // ── Electric row ──
    electricRow: {
      flexDirection: "row",
      borderLeftWidth: 1, borderRightWidth: 1, borderBottomWidth: 1,
      borderColor: BLACK, borderStyle: "solid",
    },
    eCell: {
      flex: 1,
      paddingHorizontal: sc.padH,
      paddingVertical: sc.padV,
      borderRightWidth: 1, borderColor: BLACK, borderStyle: "solid",
    },
    eCellLast: { flex: 1, paddingHorizontal: sc.padH, paddingVertical: sc.padV },
    eCellLabel: { ...B, fontSize: sc.labelSize },
    eCellValue:  { fontSize: sc.fontSize, flex: 1 },

    // ── Core Particulars table ──
    coreTable: {
      borderLeftWidth: 1, borderRightWidth: 1,
      borderTopWidth: 1, borderBottomWidth: 0,
      borderColor: BLACK, borderStyle: "solid",
      marginTop: 4,
    },
    coreHeaderRow: {
      flexDirection: "row",
      backgroundColor: "#cccccc",
      borderBottomWidth: 1, borderColor: BLACK, borderStyle: "solid",
    },
    coreRow: {
      flexDirection: "row",
      borderBottomWidth: 1, borderColor: BLACK, borderStyle: "solid",
      minHeight: sc.rowHeight,
    },
    coreSepRow: {
      flexDirection: "row",
      borderBottomWidth: 1, borderColor: BLACK, borderStyle: "solid",
      minHeight: sc.sepHeight,
    },
    colLabel: {
      width: "34%",
      paddingHorizontal: sc.padH,
      paddingVertical: sc.padV,
      borderRightWidth: 1, borderColor: BLACK, borderStyle: "solid",
      fontSize: sc.labelSize,
    },
    colCore: {
      flex: 1,
      paddingHorizontal: sc.padH,
      paddingVertical: sc.padV,
      borderRightWidth: 1, borderColor: BLACK, borderStyle: "solid",
      fontSize: sc.fontSize,
    },
    colCoreLast: {
      flex: 1,
      paddingHorizontal: sc.padH,
      paddingVertical: sc.padV,
      fontSize: sc.fontSize,
    },
    colSpanned: {
      flex: 3,
      paddingHorizontal: sc.padH,
      paddingVertical: sc.padV,
      fontSize: sc.fontSize,
    },
    coreHeaderCell: {
      paddingHorizontal: sc.padH,
      paddingVertical: sc.padV,
      fontSize: sc.labelSize,
      backgroundColor: "#cccccc",
    },

    // ── Notes & Revision ──
    noteTable: {
      borderLeftWidth: 1, borderRightWidth: 1,
      borderBottomWidth: 1, borderTopWidth: 1,
      borderColor: BLACK, borderStyle: "solid",
      marginTop: 6, flexDirection: "row",
      minHeight: 36,
    },
    noteCell: {
      flex: 4,
      paddingHorizontal: sc.padH + 1,
      paddingVertical: sc.padV,
      borderRightWidth: 1, borderColor: BLACK, borderStyle: "solid",
    },
    revCell: { flex: 1, paddingHorizontal: sc.padH + 1, paddingVertical: sc.padV },
    noteLabel: { ...B, fontSize: sc.labelSize, marginBottom: 2 },
    noteValue:  { fontSize: sc.fontSize, lineHeight: 1.3 },

    // ── Signatures ──
    sigRow:  { flexDirection: "row", marginTop: 18, paddingTop: 4 },
    sigCell: { flex: 1, alignItems: "center", paddingHorizontal: sc.padH },
    sigName: { fontSize: sc.fontSize + 1, marginBottom: 18, ...B },
    sigLine: {
      borderTopWidth: 1, borderColor: BLACK, borderStyle: "solid",
      width: "85%", marginTop: 2, paddingTop: 3,
    },
    sigLabel: { fontSize: sc.fontSize, ...B, textAlign: "center" },

    bold:   B,
    center: { textAlign: "center" },
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// COMPONENT
// ─────────────────────────────────────────────────────────────────────────────

type Props = {
  data: TiRecordInput & { ti_no?: string | null };
};

export function TiPdfDocument({ data }: Props) {
  const c1 = (data.core1 ?? {}) as CoreData;
  const c2 = (data.core2 ?? {}) as CoreData;
  const c3 = (data.core3 ?? {}) as CoreData;

  const maxExcVoltage = c1.max_exc_voltage || c2.max_exc_voltage || c3.max_exc_voltage;
  const maxExcLabel   = `Max. Exc. C/n. :- @${maxExcVoltage || "VK/2"}`;

  const rows     = buildRows(data, maxExcLabel);
  const sepRows  = rows.filter(r => r.type === "separator").length;
  const dataRows = rows.length - sepRows;

  // ── Auto-scale resolution ──────────────────────────────────────────────────
  const sc = computeScale(dataRows, sepRows);
  const s  = makeStyles(sc);

  // ─────────────────────────────────────────────────────────────────────────
  return (
    <Document>
      <Page size="A4" style={s.page}>

        {/* ── Header ──────────────────────────────────────────────────── */}
        <View style={s.headerRow}>
          <View style={s.headerLeft}>
            <Text style={s.headerTitle}>TECHNICAL INSTRUCTION</Text>
            <Text style={s.headerSub}>CURRENT TRANSFORMER</Text>
          </View>
          <View style={s.headerRight}>
            <View style={s.headerField}>
              <Text style={s.headerLabel}>TI no :</Text>
              <Text>{v(data.ti_no)}</Text>
            </View>
            <View style={s.headerField}>
              <Text style={s.headerLabel}>TI DATE :</Text>
              <Text>{formatDate(data.ti_date)}</Text>
            </View>
          </View>
        </View>

        {/* ── Info / Customer ─────────────────────────────────────────── */}
        {/* Each cell is now a single inline row: LABEL : value            */}
        <View style={s.infoTable}>
          {[
            [
              { label: "CUSTOMER NAME",          val: v(data.customer_name),         last: false },
              { label: "CUS. ORDER. NO.",         val: v(data.cus_order_no),          last: true  },
            ],
            [
              { label: "CUS. ORDER DATE",         val: formatDate(data.cus_order_date), last: false },
              { label: "Cust. Item No / Part code", val: v(data.cust_part_code),      last: true  },
            ],
            [
              { label: "W.O. NO.",                val: v(data.wo_number),             last: false },
              { label: "PO ITEM NO.",              val: v(data.po_item_no),            last: true  },
            ],
            [
              { label: "ITEM NO",                 val: v(data.item_no),               last: false },
              { label: "CT TYPE",                 val: v(data.ct_type),               last: true  },
            ],
            [
              { label: "QTY",                     val: data.quantity ? `${v(data.quantity)} NOS` : "", last: false },
              { label: "Sr. No.",                 val: v(data.serial_number),         last: true  },
            ],
          ].map((rowCells, ri, arr) => {
            const isLastRow = ri === arr.length - 1;
            return (
              <View key={ri} style={isLastRow ? s.infoRowLast : s.infoRow}>
                {rowCells.map((cell) => (
                  <View key={cell.label} style={cell.last ? s.infoCellLast : s.infoCell}>
                    <View style={s.infoInlineRow}>
                      <Text style={s.infoLabel}>{cell.label}</Text>
                      <Text style={s.infoColon}> :</Text>
                      <Text style={s.infoValue}> {cell.val}</Text>
                    </View>
                  </View>
                ))}
              </View>
            );
          })}
        </View>

        {/* ── Electrical details ──────────────────────────────────────── */}
        <View style={s.electricRow}>
          {[
            { label: "RATIO",         value: v(data.ratio),             last: false },
            { label: "RATED VOLTAGE", value: v(data.rated_voltage),     last: false },
            { label: "STC",           value: v(data.stc),               last: false },
            { label: "I.L.",          value: v(data.insulation_level),  last: false },
            { label: "FREQ.",         value: v(data.frequency),         last: false },
            { label: "REF. STD.",     value: v(data.ref_std),           last: true  },
          ].map((f) => (
            <View key={f.label} style={f.last ? s.eCellLast : s.eCell}>
              <View style={s.infoInlineRow}>
                <Text style={s.eCellLabel}>{f.label}</Text>
                <Text style={s.infoColon}> :</Text>
                <Text style={s.eCellValue}> {f.value}</Text>
              </View>
            </View>
          ))}
        </View>

        {/* ── Core Particulars table ──────────────────────────────────── */}
        <View style={s.coreTable}>

          {/* Header row */}
          <View style={s.coreHeaderRow}>
            <View style={[s.colLabel, s.coreHeaderCell]}>
              <Text style={[s.bold, s.center]}>PARTICULARS</Text>
            </View>
            <View style={[s.colCore, s.coreHeaderCell]}>
              <Text style={[s.bold, s.center]}>Core 1</Text>
            </View>
            <View style={[s.colCore, s.coreHeaderCell]}>
              <Text style={[s.bold, s.center]}>Core 2</Text>
            </View>
            <View style={[s.colCoreLast, s.coreHeaderCell]}>
              <Text style={[s.bold, s.center]}>Core 3</Text>
            </View>
          </View>

          {/* Data rows */}
          {rows.map((row, idx) => {
            if (row.type === "separator") {
              return (
                <View key={`sep-${idx}`} style={s.coreSepRow}>
                  <View style={s.colLabel}><Text> </Text></View>
                  <View style={s.colCore}><Text> </Text></View>
                  <View style={s.colCore}><Text> </Text></View>
                  <View style={s.colCoreLast}><Text> </Text></View>
                </View>
              );
            }

            if (row.type === "core") {
              return (
                <View key={idx} style={s.coreRow}>
                  <View style={s.colLabel}><Text>{row.label}</Text></View>
                  <View style={s.colCore}><Text>{v(c1[row.key])}</Text></View>
                  <View style={s.colCore}><Text>{v(c2[row.key])}</Text></View>
                  <View style={s.colCoreLast}><Text>{v(c3[row.key])}</Text></View>
                </View>
              );
            }

            // single (spans all core columns)
            return (
              <View key={idx} style={s.coreRow}>
                <View style={s.colLabel}><Text>{row.label}</Text></View>
                <View style={s.colSpanned}><Text>{row.value}</Text></View>
              </View>
            );
          })}
        </View>

        {/* ── Notes & Revision ────────────────────────────────────────── */}
        <View style={s.noteTable}>
          <View style={s.noteCell}>
            <Text style={s.noteLabel}>NOTE</Text>
            <Text style={s.noteValue}>{v(data.note)}</Text>
          </View>
          <View style={s.revCell}>
            <Text style={s.noteLabel}>REV NO.</Text>
            <Text style={s.noteValue}>{v(data.rev_no)}</Text>
          </View>
        </View>

        {/* ── Signatures ──────────────────────────────────────────────── */}
        <View style={s.sigRow}>
          {[
            { name: data.created_by,  label: "Created By" },
            { name: data.checked_by,  label: "Checked By" },
            { name: data.approved_by, label: "Approved By" },
          ].map(({ name, label }) => (
            <View key={label} style={s.sigCell}>
              <Text style={s.sigName}>{v(name)}</Text>
              <View style={s.sigLine}>
                <Text style={s.sigLabel}>{label}</Text>
              </View>
            </View>
          ))}
        </View>

      </Page>
    </Document>
  );
}
