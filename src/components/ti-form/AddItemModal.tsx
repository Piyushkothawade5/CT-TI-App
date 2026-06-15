import React from "react";
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogDescription } from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { Button } from "@/components/ui/button";
import { Label } from "@/components/ui/label";
import { useForm, Controller } from "react-hook-form";
import { useCreateItem, useDistinctCtTypes } from "@/api-client";
import type { ItemInput } from "@/api-client";
import { useToast } from "@/hooks/use-toast";

const CORE_FIELDS = [
  { label: "RATIO",                  key: "ratio" },
  { label: "Burden (VA)",            key: "burden_va" },
  { label: "Accuracy Class",         key: "accuracy_class" },
  { label: "ISF",                    key: "isf" },
  { label: "Min. Knee pt. volt.",    key: "min_knee_pt_volt" },
  { label: "Max. Rct @ 75°c",        key: "max_rct_75c" },
  { label: "Max. Exc. C/n @VK/2",   key: "max_exc_vk2" },
  { label: "Core Dimensions (bare)", key: "bare_core_dim" },
  { label: "Core Material",          key: "core_material" },
  { label: "Core weight (Kg)",       key: "core_weight_kg" },
  { label: "Sec. Total Turns",       key: "sec_total_turns" },
  { label: "Sec. Ter. Marking",      key: "sec_ter_marking" },
  { label: "Sec. Conductor (S1-S2)", key: "sec_cond_s1s2" },
  { label: "Sec. Turns (S1-S2)",     key: "sec_turns_s1s2" },
  { label: "Sec. Conductor (S2-S3)", key: "sec_cond_s2s3" },
  { label: "Sec. Turns (S2-S3)",     key: "sec_turns_s2s3" },
  { label: "Sec. Conductor (S3-S4)", key: "sec_cond_s3s4" },
  { label: "Sec. Turns (S3-S4)",     key: "sec_turns_s3s4" },
  { label: "Sec. Conductor (S4-S5)", key: "sec_cond_s4s5" },
  { label: "Sec. Turns (S4-S5)",     key: "sec_turns_s4s5" },
  { label: "Sec. Copper weight (kg)", key: "sec_copper_wt" },
  { label: "Finished Core Dim.",     key: "finished_core_dim" },
  { label: "Sec Connection",         key: "sec_connection" },
  { label: "Wire Length",            key: "wire_length" },
  { label: "Wire Colour",            key: "wire_colour" },
];

export function AddItemModal({ open, onOpenChange, itemNo, onSuccess }: {
  open: boolean; onOpenChange: (open: boolean) => void; itemNo: string; onSuccess: () => void;
}) {
  const { toast } = useToast();
  const createItemMutation = useCreateItem();
  const { data: distinctCtTypes = [] } = useDistinctCtTypes();

  // Clean item number: pure numeric (remove spaces, commas, dots)
  const cleanedItemNo = React.useMemo(
    () => itemNo.replace(/[\s,\.]+/g, "").replace(/[^0-9]/g, ""),
    [itemNo]
  );

  const form = useForm<ItemInput>({ defaultValues: { item_no: cleanedItemNo } });

  React.useEffect(() => {
    if (open) form.reset({ item_no: cleanedItemNo });
  }, [form, cleanedItemNo, open]);

  const handleSave = async () => {
    try {
      await createItemMutation.mutateAsync({ data: form.getValues() });
      toast({ title: "Item added successfully", className: "bg-green-50 border-green-200 text-green-800" });
      onSuccess();
    } catch {
      toast({ variant: "destructive", title: "Failed to add item" });
    }
  };

  // Arrow-key navigation
  const NUM_CORE_ROWS = CORE_FIELDS.length;
  const NUM_CORE_COLS = 3;
  const handleArrowKey = (e: React.KeyboardEvent<HTMLDivElement>) => {
    const ARROWS = ["ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight", "Enter"];
    if (!ARROWS.includes(e.key)) return;
    const target = e.target as HTMLElement;
    const formRoot = target.closest("#add-item-form");
    if (!formRoot) return;
    const tag = target.tagName.toLowerCase();
    if (tag !== "input" && tag !== "textarea" && tag !== "select") return;
    const gridRow = target.dataset.gridRow !== undefined ? parseInt(target.dataset.gridRow) : null;
    const gridCol = target.dataset.gridCol !== undefined ? parseInt(target.dataset.gridCol) : null;
    if (gridRow !== null && gridCol !== null) {
      e.preventDefault();
      let nr = gridRow, nc = gridCol;
      switch (e.key) {
        case "ArrowDown": case "Enter":
          if (e.key === "Enter" && gridRow === NUM_CORE_ROWS - 1) {
            formRoot.querySelector<HTMLElement>('[data-field="ct_final_dim"]')?.focus(); return;
          }
          nr = Math.min(gridRow + 1, NUM_CORE_ROWS - 1); break;
        case "ArrowUp": nr = Math.max(gridRow - 1, 0); break;
        case "ArrowRight":
          if (gridCol < NUM_CORE_COLS - 1) nc = gridCol + 1;
          else if (gridRow < NUM_CORE_ROWS - 1) { nr = gridRow + 1; nc = 0; } break;
        case "ArrowLeft":
          if (gridCol > 0) nc = gridCol - 1;
          else if (gridRow > 0) { nr = gridRow - 1; nc = NUM_CORE_COLS - 1; } break;
      }
      formRoot.querySelector<HTMLElement>(`[data-grid-row="${nr}"][data-grid-col="${nc}"]`)?.focus();
    } else {
      const all = Array.from(formRoot.querySelectorAll<HTMLElement>("input:not([disabled]), select:not([disabled]), textarea:not([disabled])"));
      const idx = all.indexOf(target);
      if (idx === -1) return;
      const isNext = e.key === "ArrowDown" || e.key === "ArrowRight" || e.key === "Enter";
      const isPrev = e.key === "ArrowUp" || e.key === "ArrowLeft";
      if (isNext && idx < all.length - 1) { e.preventDefault(); all[idx + 1].focus(); }
      if (isPrev && idx > 0) { e.preventDefault(); all[idx - 1].focus(); }
    }
  };

  // CT Type autocomplete state
  const [ctTypeOpen, setCtTypeOpen] = React.useState(false);
  const [ctQuery, setCtQuery] = React.useState("");
  const ctRef = React.useRef<HTMLDivElement>(null);
  React.useEffect(() => {
    const h = (e: MouseEvent) => { if (ctRef.current && !ctRef.current.contains(e.target as Node)) setCtTypeOpen(false); };
    document.addEventListener("mousedown", h);
    return () => document.removeEventListener("mousedown", h);
  }, []);
  const filteredCtTypes = distinctCtTypes.filter(t => !ctQuery || t.toLowerCase().includes(ctQuery.toLowerCase()));

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-5xl max-h-[90vh] flex flex-col p-0 overflow-hidden">
        <DialogHeader className="p-6 pb-4 border-b">
          <DialogTitle className="text-xl text-[#2a4080]">Add New Item</DialogTitle>
          <DialogDescription>
            Item No. <span className="font-bold text-gray-900">{cleanedItemNo}</span> does not exist. Fill in details below.
          </DialogDescription>
        </DialogHeader>

        <div onKeyDown={handleArrowKey} className="flex-1 overflow-y-auto p-6 bg-gray-50/50">
          <div id="add-item-form" className="space-y-8 pb-6">

            <section>
              <h3 className="text-[#4a6fa5] font-bold tracking-wide text-sm mb-4 border-l-4 border-[#4a6fa5] pl-3">BASIC DETAILS</h3>
              <div className="grid grid-cols-3 gap-4">
                <div className="space-y-1">
                  <Label className="text-xs uppercase text-gray-500">Item No</Label>
                  <Input disabled value={cleanedItemNo} className="bg-gray-100" />
                </div>

                {/* CT Type — autocomplete from history */}
                <div className="space-y-1 relative" ref={ctRef}>
                  <Label className="text-xs uppercase text-gray-500">CT Type</Label>
                  <Controller name="ct_type" control={form.control} render={({ field }) => (
                    <Input {...field} value={field.value || ""} autoComplete="off"
                      className="h-9 bg-white"
                      placeholder={distinctCtTypes.length ? "Type or select..." : "Type CT type..."}
                      onFocus={() => setCtTypeOpen(true)}
                      onChange={e => { field.onChange(e.target.value); setCtQuery(e.target.value); setCtTypeOpen(true); }}
                    />
                  )} />
                  {ctTypeOpen && filteredCtTypes.length > 0 && (
                    <ul className="absolute z-50 mt-0.5 w-full bg-white border border-gray-200 rounded-md shadow-lg max-h-40 overflow-y-auto text-sm">
                      {filteredCtTypes.map(opt => (
                        <li key={opt}
                          onMouseDown={e => { e.preventDefault(); form.setValue("ct_type", opt); setCtQuery(opt); setCtTypeOpen(false); }}
                          className="px-3 py-1.5 cursor-pointer hover:bg-[#4a6fa5]/10">
                          {opt}
                        </li>
                      ))}
                    </ul>
                  )}
                </div>

                <div className="space-y-1">
                  <Label className="text-xs uppercase text-gray-500">Cust. Part Code</Label>
                  <Controller name="cust_part_code" control={form.control} render={({ field }) => (
                    <Input {...field} value={field.value || ""} className="bg-white" />
                  )} />
                </div>
                {[
                  { label: "Ratio", key: "ratio" }, { label: "Rated Voltage", key: "rated_voltage" },
                  { label: "STC", key: "stc" }, { label: "Insulation Level", key: "insulation_level" },
                  { label: "Frequency", key: "frequency" }, { label: "Ref. Std.", key: "ref_std" },
                ].map(f => (
                  <div key={f.key} className="space-y-1">
                    <Label className="text-xs uppercase text-gray-500">{f.label}</Label>
                    <Controller name={f.key as any} control={form.control} render={({ field }) => (
                      <Input {...field} value={field.value || ""} className="bg-white" />
                    )} />
                  </div>
                ))}
              </div>
            </section>

            <section>
              <h3 className="text-[#4a6fa5] font-bold tracking-wide text-sm mb-4 border-l-4 border-[#4a6fa5] pl-3">CORE PARTICULARS</h3>
              <div className="overflow-x-auto border border-[#dee2e6] rounded-md bg-white">
                <table className="w-full text-sm text-left">
                  <thead className="text-xs text-white uppercase bg-[#4a6fa5]">
                    <tr>
                      <th className="px-4 py-2 border-r border-[#dee2e6]/20 w-1/4">Particulars</th>
                      <th className="px-4 py-2 border-r border-[#dee2e6]/20">Core 1</th>
                      <th className="px-4 py-2 border-r border-[#dee2e6]/20">Core 2</th>
                      <th className="px-4 py-2">Core 3</th>
                    </tr>
                  </thead>
                  <tbody>
                    {CORE_FIELDS.map((row, idx) => (
                      <tr key={idx} className="border-b border-[#dee2e6] hover:bg-gray-50 transition-colors">
                        <td className="px-4 py-1 font-medium text-gray-900 border-r border-[#dee2e6] bg-gray-50/50">{row.label}</td>
                        {[0, 1, 2].map(col => (
                          <td key={col} className={`p-0${col < 2 ? " border-r border-[#dee2e6]" : ""}`}>
                            <Controller name={`core${col + 1}.${row.key}` as any} control={form.control} render={({ field }) => (
                              <Input {...field} value={field.value || ""} data-grid-row={idx} data-grid-col={col}
                                className="border-0 shadow-none h-8 rounded-none focus-visible:ring-1 focus-visible:ring-[#4a6fa5] bg-transparent px-3" />
                            )} />
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </section>

            <section>
              <h3 className="text-[#4a6fa5] font-bold tracking-wide text-sm mb-4 border-l-4 border-[#4a6fa5] pl-3">ADDITIONAL DETAILS</h3>
              <div className="grid grid-cols-4 gap-4">
                {[
                  { label: "CT Final Dim", key: "ct_final_dim", dataField: "ct_final_dim" },
                  { label: "GA Drg", key: "ga_drg" }, { label: "INS Class", key: "ins_class" },
                  { label: "Ref TI", key: "ref_ti" }, { label: "PRI Turns", key: "pri_turns" },
                  { label: "PRI Copper", key: "pri_copper" }, { label: "Former", key: "former" },
                  { label: "PRI Length", key: "pri_length" }, { label: "PRI Weight", key: "pri_weight" },
                  { label: "Sec. Terminal", key: "sec_terminal" }, { label: "Total Weight", key: "total_weight" },
                ].map(field => (
                  <div key={field.key} className="space-y-1">
                    <Label className="text-xs uppercase text-gray-500">{field.label}</Label>
                    <Controller name={field.key as any} control={form.control} render={({ field: f }) => (
                      <Input {...f} value={f.value || ""} data-field={(field as any).dataField} className="bg-white" />
                    )} />
                  </div>
                ))}
              </div>
            </section>
          </div>
        </div>

        <div className="flex justify-end space-x-3 p-4 border-t bg-white">
          <Button variant="outline" onClick={() => onOpenChange(false)}>Cancel</Button>
          <Button onClick={handleSave} className="bg-[#2a4080] hover:bg-[#1a2850]">Save Item</Button>
        </div>
      </DialogContent>
    </Dialog>
  );
}
