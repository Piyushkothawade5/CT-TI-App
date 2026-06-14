import React from "react";
import { Document, Page, View, Text, StyleSheet } from "@react-pdf/renderer";
import type { TiRecordInput, CoreData } from "@/api-client";

const BLACK = "#000000";

const s = StyleSheet.create({
  page: {
    fontFamily: "Helvetica",
    fontSize: 8,
    paddingTop: 18,
    paddingBottom: 18,
    paddingLeft: 22,
    paddingRight: 22,
    color: BLACK,
  },

  // ── Header ────────────────────────────────────────────────
  headerRow: {
    flexDirection: "row",
    borderWidth: 1,
    borderColor: BLACK,
    borderStyle: "solid",
    paddingHorizontal: 8,
    paddingVertical: 6,
    marginBottom: 0,
    alignItems: "center",
  },
  headerLeft: { flex: 1, alignItems: "center" },
  headerTitle: { fontSize: 11, fontFamily: "Helvetica-Bold", letterSpacing: 0.5 },
  headerSub: { fontSize: 9, marginTop: 2 },
  headerRight: { flex: 1, alignItems: "flex-end" },
  headerField: { flexDirection: "row", marginBottom: 3 },
  headerLabel: { fontFamily: "Helvetica-Bold", marginRight: 4 },

  // ── Info table ────────────────────────────────────────────
  infoTable: {
    borderLeftWidth: 1,
    borderRightWidth: 1,
    borderBottomWidth: 1,
    borderTopWidth: 0,
    borderColor: BLACK,
    borderStyle: "solid",
  },
  infoRow: {
    flexDirection: "row",
    borderBottomWidth: 1,
    borderColor: BLACK,
    borderStyle: "solid",
    minHeight: 20,
  },
  infoRowLast: {
    flexDirection: "row",
    minHeight: 20,
  },
  infoCell: {
    flex: 1,
    paddingHorizontal: 5,
    paddingVertical: 3,
    borderRightWidth: 1,
    borderColor: BLACK,
    borderStyle: "solid",
  },
  infoCellLast: {
    flex: 1,
    paddingHorizontal: 5,
    paddingVertical: 3,
  },
  infoLabel: { fontFamily: "Helvetica-Bold", fontSize: 7.5 },
  infoValue: { fontSize: 8, marginTop: 1 },

  // ── Electric row ──────────────────────────────────────────
  electricRow: {
    flexDirection: "row",
    borderLeftWidth: 1,
    borderRightWidth: 1,
    borderBottomWidth: 1,
    borderColor: BLACK,
    borderStyle: "solid",
    minHeight: 24,
  },
  eCell: {
    flex: 1,
    paddingHorizontal: 4,
    paddingVertical: 3,
    borderRightWidth: 1,
    borderColor: BLACK,
    borderStyle: "solid",
  },
  eCellLast: {
    flex: 1,
    paddingHorizontal: 4,
    paddingVertical: 3,
  },
  eCellLabel: { fontFamily: "Helvetica-Bold", fontSize: 7.5 },
  eCellValue: { fontSize: 8, marginTop: 1 },

  // ── Core Particulars table ────────────────────────────────
  coreTable: {
    borderLeftWidth: 1,
    borderRightWidth: 1,
    borderTopWidth: 1,
    borderBottomWidth: 0,
    borderColor: BLACK,
    borderStyle: "solid",
    marginTop: 4,
  },
  coreHeaderRow: {
    flexDirection: "row",
    backgroundColor: "#cccccc",
    borderBottomWidth: 1,
    borderColor: BLACK,
    borderStyle: "solid",
  },
  coreRow: {
    flexDirection: "row",
    borderBottomWidth: 1,
    borderColor: BLACK,
    borderStyle: "solid",
    minHeight: 14,
  },
  colLabel: {
    width: "34%",
    paddingHorizontal: 4,
    paddingVertical: 2,
    borderRightWidth: 1,
    borderColor: BLACK,
    borderStyle: "solid",
  },
  colCore: {
    flex: 1,
    paddingHorizontal: 4,
    paddingVertical: 2,
    borderRightWidth: 1,
    borderColor: BLACK,
    borderStyle: "solid",
  },
  colCoreLast: {
    flex: 1,
    paddingHorizontal: 4,
    paddingVertical: 2,
  },
  colSpanned: {
    flex: 3,
    paddingHorizontal: 4,
    paddingVertical: 2,
  },

  // ── Signatures ────────────────────────────────────────────
  sigRow: {
    flexDirection: "row",
    marginTop: 6,
    borderTopWidth: 1,
    borderColor: BLACK,
    borderStyle: "solid",
    paddingTop: 4,
  },
  sigCell: { flex: 1 },

  bold: { fontFamily: "Helvetica-Bold" },
  center: { textAlign: "center" },
});

function v(val?: string | null): string {
  return val ?? "";
}

function formatDate(dateStr?: string | null): string {
  if (!dateStr) return "";
  try {
    const d = new Date(dateStr + "T00:00:00");
    const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
    return `${String(d.getDate()).padStart(2, "0")}-${months[d.getMonth()]}-${d.getFullYear()}`;
  } catch {
    return dateStr;
  }
}

type CoreRow =
  | { type: "core"; label: string; key: keyof CoreData }
  | { type: "single"; label: string; value: string }
  | { type: "separator" };

type Props = {
  data: TiRecordInput & { ti_no?: string | null };
};

export function TiPdfDocument({ data }: Props) {
  const c1 = data.core1 ?? {};
  const c2 = data.core2 ?? {};
  const c3 = data.core3 ?? {};

  const maxExcVoltage =
    (c1 as CoreData).max_exc_voltage ||
    (c2 as CoreData).max_exc_voltage ||
    (c3 as CoreData).max_exc_voltage;
  const maxExcLabel = `Max. Exc. C/n. :- @${maxExcVoltage || "VK/2"}`;

  const rows: CoreRow[] = [
    { type: "core", label: "RATIO", key: "ratio" },
    { type: "core", label: "Burden (VA)", key: "burden_va" },
    { type: "core", label: "Accuracy Class", key: "accuracy_class" },
    { type: "core", label: "ISF", key: "isf" },
    { type: "core", label: "Min. Knee pt. volt.", key: "min_knee_pt_volt" },
    { type: "core", label: "Max. Rct @ 75\u00b0c", key: "max_rct_75c" },
    { type: "core", label: maxExcLabel, key: "max_exc_vk2" },
    { type: "separator" },
    { type: "core", label: "Core Dimensions", key: "bare_core_dim" },
    { type: "core", label: "Core Material", key: "core_material" },
    { type: "core", label: "Core weight (Kg)", key: "core_weight_kg" },
    { type: "core", label: "Sec. Total Turns", key: "sec_total_turns" },
    { type: "core", label: "Sec. Ter. Marking", key: "sec_ter_marking" },
    { type: "core", label: "Sec. Conductor (S1-S2)", key: "sec_cond_s1s2" },
    { type: "core", label: "Sec. Turns (S1-S2)", key: "sec_turns_s1s2" },
    { type: "core", label: "Sec. Conductor (S2-S3)", key: "sec_cond_s2s3" },
    { type: "core", label: "Sec. Turns (S2-S3)", key: "sec_turns_s2s3" },
    { type: "core", label: "Sec. Conductor (S3-S4)", key: "sec_cond_s3s4" },
    { type: "core", label: "Sec. Turns (S3-S4)", key: "sec_turns_s3s4" },
    { type: "core", label: "Sec. Conductor (S4-S5)", key: "sec_cond_s4s5" },
    { type: "core", label: "Sec. Turns (S4-S5)", key: "sec_turns_s4s5" },
    { type: "core", label: "Sec. Copper weight (kg)", key: "sec_copper_wt" },
    { type: "core", label: "Finished Core Dim. (mm)", key: "finished_core_dim" },
    { type: "core", label: "Sec Connection", key: "sec_connection" },
    { type: "core", label: "Wire Length", key: "wire_length" },
    { type: "core", label: "Wire Colour", key: "wire_colour" },
    { type: "single", label: "CT final dim", value: v(data.ct_final_dim) },
    { type: "single", label: "GA Drg", value: v(data.ga_drg) },
    { type: "single", label: "INS CLASS", value: v(data.ins_class) },
    { type: "single", label: "PRI Turns", value: v(data.pri_turns) },
    { type: "single", label: "PRI Copper", value: v(data.pri_copper) },
    { type: "single", label: "Former", value: v(data.former) },
    { type: "single", label: "PRI Length", value: v(data.pri_length) },
    { type: "single", label: "PRI Weight", value: v(data.pri_weight) },
    { type: "single", label: "Sec. Terminal", value: v(data.sec_terminal) },
    { type: "single", label: "Total Weight", value: v(data.total_weight) },
    { type: "single", label: "Ref TI:", value: v(data.ref_ti) },
  ];

  return (
    <Document>
      <Page size="A4" style={s.page}>

        {/* ── Header ── */}
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

        {/* ── Info / Customer section ── */}
        <View style={s.infoTable}>
          <View style={s.infoRow}>
            <View style={s.infoCell}>
              <Text style={s.infoLabel}>CUSTOMER NAME</Text>
              <Text style={s.infoValue}>{v(data.customer_name)}</Text>
            </View>
            <View style={s.infoCellLast}>
              <Text style={s.infoLabel}>CUS. ORDER. NO.</Text>
              <Text style={s.infoValue}>{v(data.cus_order_no)}</Text>
            </View>
          </View>
          <View style={s.infoRow}>
            <View style={s.infoCell}>
              <Text style={s.infoLabel}>CUS. ORDER DATE:</Text>
              <Text style={s.infoValue}>{formatDate(data.cus_order_date)}</Text>
            </View>
            <View style={s.infoCellLast}>
              <Text style={s.infoLabel}>Cust. Item No / Part code</Text>
              <Text style={s.infoValue}>{v(data.cust_part_code)}</Text>
            </View>
          </View>
          <View style={s.infoRow}>
            <View style={s.infoCell}>
              <Text style={s.infoLabel}>W.O. NO.</Text>
              <Text style={s.infoValue}>{v(data.wo_number)}</Text>
            </View>
            <View style={s.infoCellLast}>
              <Text style={s.infoLabel}>PO ITEM NO.</Text>
              <Text style={s.infoValue}>{v(data.po_item_no)}</Text>
            </View>
          </View>
          <View style={s.infoRow}>
            <View style={s.infoCell}>
              <Text style={s.infoLabel}>ITEM NO :-</Text>
              <Text style={s.infoValue}>{v(data.item_no)}</Text>
            </View>
            <View style={s.infoCellLast}>
              <Text style={s.infoLabel}>CT TYPE</Text>
              <Text style={s.infoValue}>{v(data.ct_type)}</Text>
            </View>
          </View>
          <View style={s.infoRowLast}>
            <View style={s.infoCell}>
              <Text style={s.infoLabel}>QTY :-</Text>
              <Text style={s.infoValue}>{v(data.quantity)}</Text>
            </View>
            <View style={s.infoCellLast}>
              <Text style={s.infoLabel}>Sr. No.</Text>
              <Text style={s.infoValue}>{v(data.serial_number)}</Text>
            </View>
          </View>
        </View>

        {/* ── Electric details row ── */}
        <View style={s.electricRow}>
          {[
            { label: "RATIO :-", value: v(data.ratio) },
            { label: "RATED VOLTAGE", value: v(data.rated_voltage) },
            { label: "STC", value: v(data.stc) },
            { label: "I.L.", value: v(data.insulation_level) },
            { label: "FREQ.", value: v(data.frequency) },
          ].map((f) => (
            <View key={f.label} style={s.eCell}>
              <Text style={s.eCellLabel}>{f.label}</Text>
              <Text style={s.eCellValue}>{f.value}</Text>
            </View>
          ))}
          <View style={s.eCellLast}>
            <Text style={s.eCellLabel}>REF. STD.</Text>
            <Text style={s.eCellValue}>{v(data.ref_std)}</Text>
          </View>
        </View>

        {/* ── Core Particulars table ── */}
        <View style={s.coreTable}>
          <View style={s.coreHeaderRow}>
            <View style={[s.colLabel, { backgroundColor: "#cccccc" }]}>
              <Text style={[s.bold, s.center]}>PARTICULARS</Text>
            </View>
            <View style={[s.colCore, { backgroundColor: "#cccccc" }]}>
              <Text style={[s.bold, s.center]}>Core 1</Text>
            </View>
            <View style={[s.colCore, { backgroundColor: "#cccccc" }]}>
              <Text style={[s.bold, s.center]}>Core 2</Text>
            </View>
            <View style={[s.colCoreLast, { backgroundColor: "#cccccc" }]}>
              <Text style={[s.bold, s.center]}>Core 3</Text>
            </View>
          </View>

          {rows.map((row, idx) => {
            if (row.type === "separator") {
              return (
                <View key={`sep-${idx}`} style={[s.coreRow, { minHeight: 6 }]}>
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
                  <View style={s.colCore}>
                    <Text>{v((c1 as CoreData)[row.key])}</Text>
                  </View>
                  <View style={s.colCore}>
                    <Text>{v((c2 as CoreData)[row.key])}</Text>
                  </View>
                  <View style={s.colCoreLast}>
                    <Text>{v((c3 as CoreData)[row.key])}</Text>
                  </View>
                </View>
              );
            }
            return (
              <View key={idx} style={s.coreRow}>
                <View style={s.colLabel}><Text>{row.label}</Text></View>
                <View style={s.colSpanned}><Text>{row.value}</Text></View>
              </View>
            );
          })}
        </View>

        {/* ── Signatures ── */}
        <View style={s.sigRow}>
          <View style={s.sigCell}>
            <Text>
              <Text style={s.bold}>Created By :- </Text>
              {v(data.created_by)}
            </Text>
          </View>
          <View style={s.sigCell}>
            <Text>
              <Text style={s.bold}>Checked By :- </Text>
              {v(data.checked_by)}
            </Text>
          </View>
          <View style={s.sigCell}>
            <Text>
              <Text style={s.bold}>Approved By :- </Text>
              {v(data.approved_by)}
            </Text>
          </View>
        </View>

      </Page>
    </Document>
  );
}
