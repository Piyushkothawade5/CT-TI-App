import React, { useState, useEffect } from "react";
import { useForm, Controller } from "react-hook-form";
import {
  Save, FilePlus, Search, ChevronLeft, ChevronRight, Edit3, Printer, FileText, Settings, X, PlusCircle
} from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { useToast } from "@/hooks/use-toast";
import {
  useGetItem, useGetTiRecord, useGenerateTiNumber, useGetAdjacentTiRecords, useUpdateTiRecord, useCreateTiRecord, getGetItemQueryKey, getGetTiRecordQueryKey
} from "@/api-client";
import type { TiRecordInput, CoreData } from "@/api-client";
import { useQueryClient } from "@tanstack/react-query";
import { SearchModal } from "@/components/ti-form/SearchModal";
import { AddItemModal } from "@/components/ti-form/AddItemModal";
import { downloadTiPdf, printTiPdf } from "@/components/ti-form/downloadTiPdf";

const CORE_FIELDS: Array<{
  label: string;
  key: string;
  companionKey?: string;
  companionPlaceholder?: string;
}> = [
  { label: "RATIO", key: "ratio" },
  { label: "Burden (VA)", key: "burden_va" },
  { label: "Accuracy Class", key: "accuracy_class" },
  { label: "ISF", key: "isf" },
  { label: "Min. Knee pt. volt.", key: "min_knee_pt_volt" },
  { label: "Max. Rct @ 75°c", key: "max_rct_75c" },
  { label: "Max. Exc. C/n", key: "max_exc_vk2", companionKey: "max_exc_voltage", companionPlaceholder: "@ voltage (e.g. VK/2)" },
  { label: "Core Dimensions (bare)", key: "bare_core_dim" },
  { label: "Core Material", key: "core_material" },
  { label: "Core weight (Kg)", key: "core_weight_kg" },
  { label: "Sec. Total Turns", key: "sec_total_turns" },
  { label: "Sec. Ter. Marking", key: "sec_ter_marking" },
  { label: "Sec. Conductor (S1-S2)", key: "sec_cond_s1s2" },
  { label: "Sec. Turns (S1-S2)", key: "sec_turns_s1s2" },
  { label: "Sec. Conductor (S2-S3)", key: "sec_cond_s2s3" },
  { label: "Sec. Turns (S2-S3)", key: "sec_turns_s2s3" },
  { label: "Sec. Conductor (S3-S4)", key: "sec_cond_s3s4" },
  { label: "Sec. Turns (S3-S4)", key: "sec_turns_s3s4" },
  { label: "Sec. Conductor (S4-S5)", key: "sec_cond_s4s5" },
  { label: "Sec. Turns (S4-S5)", key: "sec_turns_s4s5" },
  { label: "Sec. Copper weight (kg)", key: "sec_copper_wt" },
  { label: "Finished Core Dim.", key: "finished_core_dim" },
  { label: "Sec Connection", key: "sec_connection" },
  { label: "Wire Length", key: "wire_length" },
  { label: "Wire Colour", key: "wire_colour" },
];

export default function Home() {
  const { toast } = useToast();
  const queryClient = useQueryClient();

  // Sticky signature fields — persist across New, carry forward to next TI
  const stickyRef = React.useRef({ approved_by: "", checked_by: "", created_by: "" });

  // Draft TI number — auto-generated on New, user-editable before Save
  const [draftTiNo, setDraftTiNo] = React.useState("");

  // Generate a TI number preview whenever we enter new mode
  const generatePreviewTiNo = React.useCallback(async () => {
    try {
      const result = await generateTiMutation.mutateAsync({});
      setDraftTiNo(result.ti_no);
    } catch {
      const now = new Date();
      const m = now.getMonth() + 1;
      const y = now.getFullYear() % 100;
      const fyS = m >= 4 ? y : y - 1;
      const fyE = fyS + 1;
      setDraftTiNo(`LTCT-${fyS}-${fyE}-????`);
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Auto-generate on initial mount (app starts in new mode)
  React.useEffect(() => {
    generatePreviewTiNo();
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // App State
  const [currentTiNo, setCurrentTiNo] = useState<string | null>(null);
  const [isEditMode, setIsEditMode] = useState(false);
  const [isNewMode, setIsNewMode] = useState(true);

  // Debounced TI number for duplicate checking (must be after isNewMode)
  const [debouncedTiNo, setDebouncedTiNo] = React.useState("");
  React.useEffect(() => {
    if (!isNewMode || !draftTiNo) { setDebouncedTiNo(""); return; }
    const t = setTimeout(() => setDebouncedTiNo(draftTiNo), 400);
    return () => clearTimeout(t);
  }, [draftTiNo, isNewMode]);
  
  const [itemNoInput, setItemNoInput] = useState("");
  const [activeItemNo, setActiveItemNo] = useState("");

  const [isSearchModalOpen, setIsSearchModalOpen] = useState(false);
  const [isAddItemModalOpen, setIsAddItemModalOpen] = useState(false);

  // Data Fetching
  const { data: itemData, isError: isItemError } = useGetItem(activeItemNo, {
    query: {
      enabled: !!activeItemNo,
      retry: false,
    }
  });

  const { data: tiRecordData } = useGetTiRecord(currentTiNo || "", {
    query: {
      enabled: !!currentTiNo,
      retry: false,
    }
  });

  const { data: adjacentData } = useGetAdjacentTiRecords(currentTiNo || "", {
    query: {
      enabled: !!currentTiNo,
    }
  });

  // Duplicate check — query the debounced draft TI number
  const { data: duplicateTiData } = useGetTiRecord(debouncedTiNo, {
    query: {
      enabled: !!debouncedTiNo && isNewMode,
      retry: false,
    }
  });
  const isDuplicateTiNo = !!duplicateTiData;

  const generateTiMutation = useGenerateTiNumber();
  const createTiMutation = useCreateTiRecord();
  const updateTiMutation = useUpdateTiRecord();

  const form = useForm<TiRecordInput>({
    defaultValues: {
      ti_date: new Date().toISOString().split("T")[0],
      item_no: "",
      approved_by: "",
      checked_by: "",
      created_by: "",
    }
  });

  // Effect: When Item data loads
  useEffect(() => {
    if (itemData && isNewMode) {
      toast({
        title: "Item loaded from database",
        className: "bg-green-50 border-green-200 text-green-800",
      });
      // Populate fields from itemData
      const currentValues = form.getValues();
      form.reset({
        ...currentValues,
        item_no: itemData.item_no,
        ct_type: itemData.ct_type,
        cust_part_code: itemData.cust_part_code,
        ratio: itemData.ratio,
        rated_voltage: itemData.rated_voltage,
        stc: itemData.stc,
        insulation_level: itemData.insulation_level,
        frequency: itemData.frequency,
        ref_std: itemData.ref_std,
        core1: itemData.core1 || {},
        core2: itemData.core2 || {},
        core3: itemData.core3 || {},
        ct_final_dim: itemData.ct_final_dim,
        ga_drg: itemData.ga_drg,
        ins_class: itemData.ins_class,
        pri_turns: itemData.pri_turns,
        pri_copper: itemData.pri_copper,
        former: itemData.former,
        pri_length: itemData.pri_length,
        pri_weight: itemData.pri_weight,
        sec_terminal: itemData.sec_terminal,
        total_weight: itemData.total_weight,
        ref_ti: itemData.ref_ti,
      });
    }
  }, [itemData]);

  // Effect: Item Not Found
  useEffect(() => {
    if (isItemError && activeItemNo) {
      setIsAddItemModalOpen(true);
    }
  }, [isItemError, activeItemNo]);

  // Effect: When TI Record loads — also update sticky signature fields
  useEffect(() => {
    if (tiRecordData) {
      setIsNewMode(false);
      setIsEditMode(false);
      setItemNoInput(tiRecordData.item_no);
      setActiveItemNo(tiRecordData.item_no);
      form.reset(tiRecordData);
      // Update sticky ref from loaded record
      stickyRef.current = {
        approved_by: tiRecordData.approved_by || "",
        checked_by: tiRecordData.checked_by || "",
        created_by: tiRecordData.created_by || "",
      };
    }
  }, [tiRecordData]);

  const handleItemSearch = () => {
    if (!itemNoInput) return;
    setActiveItemNo(itemNoInput);
  };

  const isFormEnabled = !!activeItemNo && !isItemError && (isNewMode || isEditMode);

  // Actions
  const handleNew = async () => {
    setCurrentTiNo(null);
    setDraftTiNo("");
    setIsNewMode(true);
    setIsEditMode(true);
    setItemNoInput("");
    setActiveItemNo("");
    form.reset({
      ti_date: new Date().toISOString().split("T")[0],
      item_no: "",
      approved_by: stickyRef.current.approved_by,
      checked_by: stickyRef.current.checked_by,
      created_by: stickyRef.current.created_by,
    });
    // Auto-generate a preview TI number
    try {
      const result = await generateTiMutation.mutateAsync({});
      setDraftTiNo(result.ti_no);
    } catch {
      // fallback: calculate client-side
      const now = new Date();
      const m = now.getMonth() + 1;
      const y = now.getFullYear() % 100;
      const fyS = m >= 4 ? y : y - 1;
      const fyE = fyS + 1;
      setDraftTiNo(`LTCT-${String(fyS).padStart(2,"0")}-${String(fyE).padStart(2,"0")}-????`);
    }
  };

  const handleEdit = () => {
    if (currentTiNo) {
      setIsEditMode(true);
    }
  };

  const handleSave = async () => {
    if (!isFormEnabled) return;
    if (isDuplicateTiNo) {
      toast({ variant: "destructive", title: "TI number already exists", description: `${draftTiNo} is already in use. Change it before saving.` });
      return;
    }
    
    const data = form.getValues();
    if (!data.item_no) {
      toast({ variant: "destructive", title: "Item number is required" });
      return;
    }

    try {
      if (isNewMode) {
        const res = await createTiMutation.mutateAsync({ data: { ...data, ti_no: draftTiNo || undefined } });
        setCurrentTiNo(res.ti_no);
        setDraftTiNo("");
        setIsNewMode(false);
        setIsEditMode(false);
        // Capture sticky signature fields after save
        stickyRef.current = {
          approved_by: data.approved_by || "",
          checked_by: data.checked_by || "",
          created_by: data.created_by || "",
        };
        toast({ title: "Record saved successfully" });
      } else if (currentTiNo) {
        await updateTiMutation.mutateAsync({ tiNo: currentTiNo, data });
        setIsEditMode(false);
        // Update sticky fields in case they were edited
        stickyRef.current = {
          approved_by: data.approved_by || "",
          checked_by: data.checked_by || "",
          created_by: data.created_by || "",
        };
        toast({ title: "Record updated successfully" });
      }
    } catch (e) {
      toast({ variant: "destructive", title: "Error saving record" });
    }
  };

  // ── Arrow-key navigation ───────────────────────────────
  const NUM_CORE_ROWS = CORE_FIELDS.length; // 25
  const NUM_CORE_COLS = 3;

  const handleArrowKey = (e: React.KeyboardEvent<HTMLDivElement>) => {
    const ARROWS = ["ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight", "Enter"];
    if (!ARROWS.includes(e.key)) return;

    const target = e.target as HTMLElement;
    const tag = target.tagName.toLowerCase();
    if (tag !== "input" && tag !== "textarea" && tag !== "select") return;

    const gridRow = target.dataset.gridRow !== undefined ? parseInt(target.dataset.gridRow) : null;
    const gridCol = target.dataset.gridCol !== undefined ? parseInt(target.dataset.gridCol) : null;

    if (gridRow !== null && gridCol !== null) {
      // ── 2-D navigation inside the Core Particulars table ──
      e.preventDefault();
      let nr = gridRow, nc = gridCol;
      switch (e.key) {
        case "ArrowDown":
        case "Enter":
          // Special case: if on last row (Wire Colour) and ENTER is pressed, jump to CT Final Dim
          if (e.key === "Enter" && gridRow === NUM_CORE_ROWS - 1) {
            const ctFinalDim = document.querySelector<HTMLElement>('[data-field="ct_final_dim"]');
            ctFinalDim?.focus();
            return;
          }
          nr = Math.min(gridRow + 1, NUM_CORE_ROWS - 1);
          break;
        case "ArrowUp":
          nr = Math.max(gridRow - 1, 0);
          break;
        case "ArrowRight":
          if (gridCol < NUM_CORE_COLS - 1) { nc = gridCol + 1; }
          else if (gridRow < NUM_CORE_ROWS - 1) { nr = gridRow + 1; nc = 0; }
          break;
        case "ArrowLeft":
          if (gridCol > 0) { nc = gridCol - 1; }
          else if (gridRow > 0) { nr = gridRow - 1; nc = NUM_CORE_COLS - 1; }
          break;
      }
      const next = document.querySelector<HTMLElement>(
        `[data-grid-row="${nr}"][data-grid-col="${nc}"]`
      );
      next?.focus();
    } else {
      // ── Linear navigation for all other fields ──────────
      // Left/Right behave the same as Up/Down outside the table
      const form = target.closest("#ti-form");
      if (!form) return;
      const all = Array.from(
        form.querySelectorAll<HTMLElement>(
          "input:not([disabled]), select:not([disabled]), textarea:not([disabled])"
        )
      );
      const idx = all.indexOf(target);
      if (idx === -1) return;
      const isNext = e.key === "ArrowDown" || e.key === "ArrowRight" || e.key === "Enter";
      const isPrev = e.key === "ArrowUp"   || e.key === "ArrowLeft";
      if (isNext && idx < all.length - 1) { e.preventDefault(); all[idx + 1].focus(); }
      if (isPrev && idx > 0)              { e.preventDefault(); all[idx - 1].focus(); }
    }
  };
  // ────────────────────────────────────────────────────────

  const handlePrev = () => {
    if (adjacentData?.prev) {
      setCurrentTiNo(adjacentData.prev);
    } else {
      toast({ title: "No previous record" });
    }
  };

  const handleNext = () => {
    if (adjacentData?.next) {
      setCurrentTiNo(adjacentData.next);
    } else {
      toast({ title: "No next record" });
    }
  };

  const handleDownloadPdf = async () => {
    const data = form.getValues();
    const tiNo = currentTiNo || draftTiNo || "";
    try {
      await downloadTiPdf({ ...data, ti_no: tiNo });
    } catch (err) {
      toast({ title: "PDF generation failed", description: String(err), variant: "destructive" });
    }
  };

  const handlePrintPdf = async () => {
    const data = form.getValues();
    const tiNo = currentTiNo || draftTiNo || "";
    try {
      await printTiPdf({ ...data, ti_no: tiNo });
    } catch (err) {
      toast({ title: "PDF generation failed", description: String(err), variant: "destructive" });
    }
  };

  return (
    <div className="flex min-h-screen bg-gray-100">
      {/* Sidebar Toolbar */}
      <div className="w-[60px] bg-[#2a4080] flex flex-col items-center py-4 space-y-4 no-print shrink-0 fixed h-full z-10">
        <SidebarButton icon={<Save />} title="Save" onClick={handleSave} disabled={!isFormEnabled} />
        <SidebarButton icon={<FilePlus />} title="New" onClick={handleNew} />
        <SidebarButton icon={<Search />} title="Search" onClick={() => setIsSearchModalOpen(true)} />
        <SidebarButton icon={<ChevronLeft />} title="Prev" onClick={handlePrev} />
        <SidebarButton icon={<ChevronRight />} title="Next" onClick={handleNext} />
        <SidebarButton icon={<Edit3 />} title="Edit" onClick={handleEdit} disabled={!currentTiNo || isEditMode} />
        <SidebarButton icon={<Printer />} title="Print" onClick={handlePrintPdf} />
        <SidebarButton icon={<FileText />} title="PDF" onClick={handleDownloadPdf} />
      </div>

      {/* Main Form Area */}
      <div id="ti-form" onKeyDown={handleArrowKey} className="ml-[60px] flex-1 p-6 flex justify-center">
        <div className="w-full max-w-5xl bg-white shadow-lg border border-gray-200 print:shadow-none print:border-none print:max-w-none">
          {/* Header */}
          <div className="bg-gradient-to-r from-[#3b5fc0] to-[#6b8dd6] p-6 text-white flex justify-between items-center print-break-inside-avoid">
            <div className="flex items-center space-x-4">
              <div className="w-12 h-12 bg-white rounded-full flex items-center justify-center text-[#3b5fc0] font-bold text-xl shadow">
                <Settings />
              </div>
              <div>
                <h1 className="text-xl font-bold tracking-wider">TECHNICAL INSTRUCTION</h1>
                <h2 className="text-sm font-medium text-blue-100">CURRENT TRANSFORMER</h2>
              </div>
            </div>
            <div className="text-right space-y-2">
              <div className={`flex items-center space-x-2 px-3 py-1 rounded transition-colors ${isDuplicateTiNo ? "bg-red-500/80" : "bg-white/20"}`}>
                <span className="text-sm font-semibold whitespace-nowrap">TI No:</span>
                {isNewMode ? (
                  <input
                    type="text"
                    value={draftTiNo}
                    onChange={(e) => setDraftTiNo(e.target.value)}
                    placeholder="Auto-generating..."
                    className="font-mono font-bold tracking-wider bg-transparent border-b border-white/60 outline-none text-white placeholder:text-white/50 w-44 text-right"
                  />
                ) : (
                  <span className="font-mono font-bold tracking-wider">{currentTiNo || "—"}</span>
                )}
              </div>
              {isDuplicateTiNo && (
                <p className="text-xs text-red-200 font-medium text-right pr-1">⚠ This TI number already exists</p>
              )}
              <div className="flex items-center space-x-2 justify-end">
                <span className="text-sm font-semibold">TI DATE:</span>
                <Controller
                  name="ti_date"
                  control={form.control}
                  render={({ field }) => (
                    <input 
                      type="date" 
                      className="bg-transparent border border-white/30 rounded px-2 py-1 text-sm outline-none text-white [&::-webkit-calendar-picker-indicator]:filter-white disabled:opacity-80"
                      disabled={!isEditMode && !isNewMode}
                      {...field}
                    />
                  )}
                />
              </div>
            </div>
          </div>

          {/* Form Content */}
          <div className="p-6 space-y-8">
            {/* Item Number Lead */}
            <div className="bg-blue-50 border border-blue-100 p-4 rounded-md print-break-inside-avoid">
              <Label className="text-lg font-bold text-[#2a4080] mb-2 block">Item Number</Label>
              <div className="flex space-x-2">
                <Input 
                  value={itemNoInput}
                  onChange={(e) => {
                    setItemNoInput(e.target.value);
                    form.setValue("item_no", e.target.value);
                  }}
                  onBlur={handleItemSearch}
                  onKeyDown={(e) => e.key === 'Enter' && handleItemSearch()}
                  placeholder="Enter item number first..."
                  className="text-lg py-6 max-w-sm border-[#4a6fa5] focus-visible:ring-[#4a6fa5] disabled:bg-gray-200"
                  disabled={!isNewMode && !isEditMode}
                />
                <Button 
                  onClick={handleItemSearch} 
                  className="bg-[#4a6fa5] hover:bg-[#3b5fc0] h-auto px-6"
                  disabled={!isNewMode && !isEditMode}
                >
                  Load Item
                </Button>
              </div>
              {isItemError && <p className="text-red-500 text-sm mt-2 font-medium">Item not found. Please add it.</p>}
            </div>

            <div className={`space-y-8 transition-opacity duration-300 ${!isFormEnabled ? 'opacity-50 pointer-events-none grayscale-[0.2]' : ''}`}>
              
              {/* Customer Details */}
              <section className="print-break-inside-avoid">
                <SectionHeader title="CUSTOMER DETAILS" />
                <div className="grid grid-cols-2 gap-x-8 gap-y-4">
                  <FormField form={form} name="customer_name" label="Customer Name" disabled={!isFormEnabled} />
                  <FormField form={form} name="cust_part_code" label="Cust. Part Name/Item No." disabled={!isFormEnabled} />
                  <FormField form={form} name="cus_order_no" label="Customer Order No." disabled={!isFormEnabled} />
                  <FormField form={form} name="cus_order_date" label="Customer Order Date" type="date" disabled={!isFormEnabled} />
                </div>
              </section>

              {/* Order Details */}
              <section className="print-break-inside-avoid">
                <SectionHeader title="ORDER DETAILS" />
                <div className="grid grid-cols-2 gap-x-8 gap-y-4">
                  <FormField form={form} name="wo_number" label="W.O. Number" disabled={!isFormEnabled} />
                  <div className="space-y-1">
                    <Label className="text-xs font-semibold text-gray-600 uppercase tracking-wider">CT Type</Label>
                    <Controller
                      name="ct_type"
                      control={form.control}
                      render={({ field }) => (
                        <Select disabled={!isFormEnabled} value={field.value || ""} onValueChange={field.onChange}>
                          <SelectTrigger className="h-9 bg-gray-50 disabled:bg-[#f0f0f0]">
                            <SelectValue placeholder="Select CT Type" />
                          </SelectTrigger>
                          <SelectContent>
                            <SelectItem value="TAPE INSULATED CT">TAPE INSULATED CT</SelectItem>
                            <SelectItem value="PVC TAPE INSULATED CT">PVC TAPE INSULATED CT</SelectItem>
                            <SelectItem value="RESIN CAST CT">RESIN CAST CT</SelectItem>
                            <SelectItem value="RESIN INSULATED CT">RESIN INSULATED CT</SelectItem>
                            <SelectItem value="VARNISHED INSULATED CT">VARNISHED INSULATED CT</SelectItem>
                            <SelectItem value="FG TAPE INSULATED CT">FG TAPE INSULATED CT</SelectItem>
                            <SelectItem value="TAPE WOUND CT">TAPE WOUND CT</SelectItem>
                            <SelectItem value="PLASTIC CASE CT">PLASTIC CASE CT</SelectItem>
                          </SelectContent>
                        </Select>
                      )}
                    />
                  </div>
                  <FormField form={form} name="po_item_no" label="PO Item No." disabled={!isFormEnabled} />
                  <FormField form={form} name="serial_number" label="Serial Number" disabled={!isFormEnabled} />
                  <div className="space-y-1">
                    <Label className="text-xs font-semibold text-gray-600 uppercase tracking-wider">Quantity</Label>
                    <div className="flex">
                      <Controller
                        name="quantity"
                        control={form.control}
                        render={({ field }) => (
                          <Input
                            {...field}
                            value={field.value || ""}
                            disabled={!isFormEnabled}
                            className="rounded-r-none border-r-0 bg-gray-50"
                          />
                        )}
                      />
                      <div className="px-3 py-2 bg-gray-100 border border-l-0 border-gray-300 rounded-r-md text-sm text-gray-600 font-medium flex items-center">
                        NOS
                      </div>
                    </div>
                  </div>
                </div>
              </section>

              {/* Electric Details */}
              <section className="print-break-inside-avoid">
                <SectionHeader title="ELECTRIC DETAILS" />
                <div className="grid grid-cols-6 gap-4">
                  <FormField form={form} name="ratio" label="Ratio" disabled={!isFormEnabled} />
                  <FormField form={form} name="rated_voltage" label="Rated Voltage" disabled={!isFormEnabled} />
                  <FormField form={form} name="stc" label="STC" disabled={!isFormEnabled} />
                  <FormField form={form} name="insulation_level" label="I.L." disabled={!isFormEnabled} />
                  <FormField form={form} name="frequency" label="Frequency" disabled={!isFormEnabled} />
                  <FormField form={form} name="ref_std" label="Ref. Std." disabled={!isFormEnabled} />
                </div>
              </section>

              {/* Core Particulars Table */}
              <section className="print-break-inside-avoid">
                <SectionHeader title="CORE PARTICULARS" />
                <div className="overflow-x-auto border border-[#dee2e6] rounded-md">
                  <table className="w-full text-sm text-left">
                    <thead className="text-xs text-white uppercase bg-[#4a6fa5]">
                      <tr>
                        <th className="px-4 py-3 border-r border-[#dee2e6]/20 w-1/4">Particulars</th>
                        <th className="px-4 py-3 border-r border-[#dee2e6]/20">Core 1</th>
                        <th className="px-4 py-3 border-r border-[#dee2e6]/20">Core 2</th>
                        <th className="px-4 py-3">Core 3</th>
                      </tr>
                    </thead>
                    <tbody>
                      {CORE_FIELDS.map((row, idx) => (
                        <tr key={idx} className="bg-white border-b border-[#dee2e6] hover:bg-gray-50 transition-colors">
                          <td className="px-4 py-1.5 font-medium text-gray-900 border-r border-[#dee2e6] bg-gray-50/50 whitespace-nowrap">
                            {row.label}
                          </td>
                          {(["core1", "core2", "core3"] as const).map((coreKey, colIdx) => (
                            <td key={coreKey} className={`p-0${colIdx < 2 ? " border-r border-[#dee2e6]" : ""}`}>
                              {row.companionKey ? (
                                <CompanionTableCell
                                  form={form}
                                  mainName={`${coreKey}.${row.key}`}
                                  companionName={`${coreKey}.${row.companionKey}`}
                                  companionPlaceholder={row.companionPlaceholder}
                                  disabled={!isFormEnabled}
                                  gridRow={idx}
                                  gridCol={colIdx}
                                />
                              ) : (
                                <TableInput form={form} name={`${coreKey}.${row.key}`} disabled={!isFormEnabled} gridRow={idx} gridCol={colIdx} />
                              )}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </section>

              {/* Bottom Fields */}
              <section className="print-break-inside-avoid">
                <div className="grid grid-cols-4 gap-4 mt-6">
                  <FormField form={form} name="ct_final_dim" label="CT Final Dim" disabled={!isFormEnabled} dataField="ct_final_dim" />
                  <FormField form={form} name="ga_drg" label="GA Drg" disabled={!isFormEnabled} />
                  <FormField form={form} name="ins_class" label="INS Class" disabled={!isFormEnabled} />
                  <FormField form={form} name="ref_ti" label="Ref TI" disabled={!isFormEnabled} />
                  
                  <FormField form={form} name="pri_turns" label="PRI Turns" disabled={!isFormEnabled} />
                  <FormField form={form} name="pri_copper" label="PRI Copper" disabled={!isFormEnabled} />
                  <FormField form={form} name="former" label="Former" disabled={!isFormEnabled} />
                  <FormField form={form} name="pri_length" label="PRI Length" disabled={!isFormEnabled} />
                  
                  <FormField form={form} name="pri_weight" label="PRI Weight" disabled={!isFormEnabled} />
                  <FormField form={form} name="sec_terminal" label="Sec. Terminal" disabled={!isFormEnabled} />
                  <FormField form={form} name="total_weight" label="Total Weight" disabled={!isFormEnabled} />
                </div>
              </section>

              {/* Note & Rev No. */}
              <section className="print-break-inside-avoid">
                <SectionHeader title="NOTES & REVISION" />
                <div className="grid grid-cols-4 gap-4">
                  <FormField form={form} name="rev_no" label="Rev No." disabled={!isFormEnabled} />
                  <div className="col-span-3 space-y-1">
                    <label className="text-xs font-semibold text-gray-600 uppercase tracking-wider">Note</label>
                    <Controller
                      name="note"
                      control={form.control}
                      render={({ field }) => (
                        <textarea
                          {...field}
                          value={field.value || ""}
                          disabled={!isFormEnabled}
                          rows={3}
                          className="w-full px-3 py-2 text-sm border border-gray-300 rounded-md bg-gray-50 focus:outline-none focus:ring-1 focus:ring-[#4a6fa5] focus:border-[#4a6fa5] disabled:bg-[#f0f0f0] disabled:text-gray-700 disabled:opacity-100 disabled:border-transparent resize-none print:border-b print:rounded-none"
                        />
                      )}
                    />
                  </div>
                </div>
              </section>

              {/* Signatures */}
              <section className="print-break-inside-avoid pb-8">
                <SectionHeader title="SIGNATURES" />
                <p className="text-xs text-gray-500 mb-3 italic">These fields carry forward to new records. Edit here to update for future TIs.</p>
                <div className="grid grid-cols-3 gap-6">
                  {(["approved_by", "checked_by", "created_by"] as const).map((fieldName, i) => {
                    const labels = ["Approved By", "Checked By", "Created By"];
                    return (
                      <div key={fieldName} className="space-y-1">
                        <label className="text-xs font-semibold text-gray-600 uppercase tracking-wider">{labels[i]}</label>
                        <Controller
                          name={fieldName}
                          control={form.control}
                          render={({ field }) => (
                            <Input
                              {...field}
                              value={field.value || ""}
                              disabled={!isFormEnabled}
                              className="h-9 bg-gray-50 border-gray-300 focus-visible:ring-[#4a6fa5] disabled:bg-[#f0f0f0] disabled:text-gray-700 disabled:opacity-100 disabled:border-transparent"
                            />
                          )}
                        />
                        <div className="border-t border-gray-400 mt-6 pt-1 text-center text-[10px] text-gray-500 uppercase tracking-wider">{labels[i]}</div>
                      </div>
                    );
                  })}
                </div>
              </section>

            </div>
          </div>
        </div>
      </div>

      <SearchModal 
        open={isSearchModalOpen} 
        onOpenChange={setIsSearchModalOpen} 
        onSelect={(tiNo) => {
          setCurrentTiNo(tiNo);
          setIsSearchModalOpen(false);
        }}
      />
      
      <AddItemModal 
        open={isAddItemModalOpen} 
        onOpenChange={setIsAddItemModalOpen} 
        itemNo={activeItemNo}
        onSuccess={() => {
          queryClient.invalidateQueries({ queryKey: getGetItemQueryKey(activeItemNo) });
          setIsAddItemModalOpen(false);
        }}
      />

    </div>
  );
}

function SidebarButton({ icon, title, onClick, disabled }: { icon: React.ReactNode, title: string, onClick?: () => void, disabled?: boolean }) {
  return (
    <button 
      onClick={onClick}
      disabled={disabled}
      className={`w-12 h-12 flex flex-col items-center justify-center rounded-md transition-colors group
        ${disabled ? 'text-white/30 cursor-not-allowed' : 'text-white/80 hover:text-white hover:bg-white/10'}`}
    >
      <div className={`[&>svg]:w-5 [&>svg]:h-5 mb-1 transition-transform ${!disabled && 'group-hover:scale-110'}`}>
        {icon}
      </div>
      <span className="text-[10px] font-medium">{title}</span>
    </button>
  );
}

function SectionHeader({ title }: { title: string }) {
  return (
    <div className="bg-blue-50/80 border-l-4 border-[#4a6fa5] py-2 px-4 mb-4">
      <h3 className="text-[#4a6fa5] font-bold tracking-wide text-sm">{title}</h3>
    </div>
  );
}

function FormField({ form, name, label, type = "text", disabled, dataField }: { form: any, name: string, label: string, type?: string, disabled?: boolean, dataField?: string }) {
  return (
    <div className="space-y-1">
      <Label className="text-xs font-semibold text-gray-600 uppercase tracking-wider">{label}</Label>
      <Controller
        name={name}
        control={form.control}
        render={({ field }) => (
          <Input
            {...field}
            type={type}
            value={field.value || ""}
            disabled={disabled}
            data-field={dataField}
            className="h-9 bg-gray-50 border-gray-300 focus-visible:ring-[#4a6fa5] disabled:bg-[#f0f0f0] disabled:text-gray-700 disabled:opacity-100 disabled:border-transparent print:border-b print:rounded-none"
          />
        )}
      />
    </div>
  );
}

function CompanionTableCell({ form, mainName, companionName, companionPlaceholder, disabled, gridRow, gridCol }: {
  form: any;
  mainName: string;
  companionName: string;
  companionPlaceholder?: string;
  disabled?: boolean;
  gridRow?: number;
  gridCol?: number;
}) {
  return (
    <div className="flex flex-col divide-y divide-gray-100">
      <Controller
        name={mainName}
        control={form.control}
        render={({ field }) => (
          <Input
            {...field}
            value={field.value || ""}
            disabled={disabled}
            data-grid-row={gridRow}
            data-grid-col={gridCol}
            className="border-0 shadow-none h-8 rounded-none focus-visible:ring-1 focus-visible:ring-[#4a6fa5] focus-visible:ring-inset focus-visible:z-10 bg-transparent disabled:opacity-100 disabled:text-gray-900 disabled:cursor-default px-3 print:px-1"
          />
        )}
      />
      <Controller
        name={companionName}
        control={form.control}
        render={({ field }) => (
          <Input
            {...field}
            value={field.value || ""}
            disabled={disabled}
            placeholder={disabled ? "" : companionPlaceholder}
            className="border-0 shadow-none h-6 rounded-none text-[11px] text-gray-400 focus-visible:ring-1 focus-visible:ring-[#4a6fa5] focus-visible:ring-inset focus-visible:z-10 bg-transparent disabled:opacity-100 disabled:cursor-default px-3 print:px-1 placeholder:text-gray-300 italic"
          />
        )}
      />
    </div>
  );
}

function TableInput({ form, name, disabled, gridRow, gridCol }: { form: any, name: string, disabled?: boolean, gridRow?: number, gridCol?: number }) {
  return (
    <Controller
      name={name}
      control={form.control}
      render={({ field }) => (
        <Input 
          {...field} 
          value={field.value || ""}
          disabled={disabled}
          data-grid-row={gridRow}
          data-grid-col={gridCol}
          className="border-0 shadow-none h-8 rounded-none focus-visible:ring-1 focus-visible:ring-[#4a6fa5] focus-visible:ring-inset focus-visible:z-10 bg-transparent disabled:opacity-100 disabled:text-gray-900 disabled:cursor-default px-3 print:px-1" 
        />
      )}
    />
  );
}
