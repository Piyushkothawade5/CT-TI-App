import React from "react";
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogDescription } from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { Button } from "@/components/ui/button";
import { Label } from "@/components/ui/label";
import { useForm, Controller } from "react-hook-form";
import { useCreateItem } from "@/api-client";
import type { ItemInput } from "@/api-client";
import { useToast } from "@/hooks/use-toast";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";

const CORE_FIELDS = [
  { label: "RATIO", key: "ratio" },
  { label: "Burden (VA)", key: "burden_va" },
  { label: "Accuracy Class", key: "accuracy_class" },
  { label: "ISF", key: "isf" },
  { label: "Min. Knee pt. volt.", key: "min_knee_pt_volt" },
  { label: "Max. Rct @ 75°c", key: "max_rct_75c" },
  { label: "Max. Exc. C/n @VK/2", key: "max_exc_vk2" },
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

export function AddItemModal({ open, onOpenChange, itemNo, onSuccess }: { open: boolean, onOpenChange: (open: boolean) => void, itemNo: string, onSuccess: () => void }) {
  const { toast } = useToast();
  const createItemMutation = useCreateItem();

  const form = useForm<ItemInput>({
    defaultValues: {
      item_no: itemNo
    }
  });

  // Start every new item entry with a clean form.
  React.useEffect(() => {
    if (open) {
      form.reset({ item_no: itemNo });
    }
  }, [form, itemNo, open]);

  const handleSave = async () => {
    try {
      await createItemMutation.mutateAsync({ data: form.getValues() });
      toast({ title: "Item added successfully", className: "bg-green-50 border-green-200 text-green-800" });
      onSuccess();
    } catch (e) {
      toast({ variant: "destructive", title: "Failed to add item" });
    }
  };

  // ── Arrow-key navigation ───────────────────────────────
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
      // ── 2-D navigation inside the Core Particulars table ──
      e.preventDefault();
      let nr = gridRow, nc = gridCol;
      switch (e.key) {
        case "ArrowDown":
        case "Enter":
          // Special case: if on last row (Wire Colour) and ENTER is pressed, jump to CT Final Dim
          if (e.key === "Enter" && gridRow === NUM_CORE_ROWS - 1) {
            const ctFinalDim = formRoot.querySelector<HTMLElement>('[data-field="ct_final_dim"]');
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
      const next = formRoot.querySelector<HTMLElement>(
        `[data-grid-row="${nr}"][data-grid-col="${nc}"]`
      );
      next?.focus();
    } else {
      // ── Linear navigation for all other fields ──────────
      const all = Array.from(
        formRoot.querySelectorAll<HTMLElement>(
          "input:not([disabled]), select:not([disabled]), textarea:not([disabled])"
        )
      );
      const idx = all.indexOf(target);
      if (idx === -1) return;
      const isNext = e.key === "ArrowDown" || e.key === "ArrowRight" || e.key === "Enter";
      const isPrev = e.key === "ArrowUp" || e.key === "ArrowLeft";
      if (isNext && idx < all.length - 1) { e.preventDefault(); all[idx + 1].focus(); }
      if (isPrev && idx > 0) { e.preventDefault(); all[idx - 1].focus(); }
    }
  };
  // ────────────────────────────────────────────────────────

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-5xl max-h-[90vh] flex flex-col p-0 overflow-hidden">
        <DialogHeader className="p-6 pb-4 border-b">
          <DialogTitle className="text-xl text-[#2a4080]">Add New Item</DialogTitle>
          <DialogDescription>
            Item No. <span className="font-bold text-gray-900">{itemNo}</span> does not exist in the database. Please provide item details to add it to the database.
          </DialogDescription>
        </DialogHeader>

        <div onKeyDown={handleArrowKey} className="flex-1 overflow-y-auto p-6 bg-gray-50/50">
          <div id="add-item-form" className="space-y-8 pb-6">
            
            <section>
              <h3 className="text-[#4a6fa5] font-bold tracking-wide text-sm mb-4 border-l-4 border-[#4a6fa5] pl-3">BASIC DETAILS</h3>
              <div className="grid grid-cols-3 gap-4">
                <div className="space-y-1">
                  <Label className="text-xs uppercase text-gray-500">Item No</Label>
                  <Input disabled value={itemNo} className="bg-gray-100" />
                </div>
                <div className="space-y-1">
                  <Label className="text-xs uppercase text-gray-500">CT Type</Label>
                  <Controller
                    name="ct_type"
                    control={form.control}
                    render={({ field }) => (
                      <Select value={field.value || ""} onValueChange={field.onChange}>
                        <SelectTrigger className="h-9 bg-white" onKeyDown={handleArrowKey}>
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
                <div className="space-y-1">
                  <Label className="text-xs uppercase text-gray-500">Cust. Part Code</Label>
                  <Controller
                    name="cust_part_code"
                    control={form.control}
                    render={({ field }) => (
                      <Input {...field} value={field.value || ""} onKeyDown={handleArrowKey} className="bg-white" />
                    )}
                  />
                </div>
                <div className="space-y-1">
                  <Label className="text-xs uppercase text-gray-500">Ratio</Label>
                  <Controller
                    name="ratio"
                    control={form.control}
                    render={({ field }) => (
                      <Input {...field} value={field.value || ""} onKeyDown={handleArrowKey} className="bg-white" />
                    )}
                  />
                </div>
                <div className="space-y-1">
                  <Label className="text-xs uppercase text-gray-500">Rated Voltage</Label>
                  <Controller
                    name="rated_voltage"
                    control={form.control}
                    render={({ field }) => (
                      <Input {...field} value={field.value || ""} onKeyDown={handleArrowKey} className="bg-white" />
                    )}
                  />
                </div>
                <div className="space-y-1">
                  <Label className="text-xs uppercase text-gray-500">STC</Label>
                  <Controller
                    name="stc"
                    control={form.control}
                    render={({ field }) => (
                      <Input {...field} value={field.value || ""} onKeyDown={handleArrowKey} className="bg-white" />
                    )}
                  />
                </div>
                <div className="space-y-1">
                  <Label className="text-xs uppercase text-gray-500">Insulation Level</Label>
                  <Controller
                    name="insulation_level"
                    control={form.control}
                    render={({ field }) => (
                      <Input {...field} value={field.value || ""} onKeyDown={handleArrowKey} className="bg-white" />
                    )}
                  />
                </div>
                <div className="space-y-1">
                  <Label className="text-xs uppercase text-gray-500">Frequency</Label>
                  <Controller
                    name="frequency"
                    control={form.control}
                    render={({ field }) => (
                      <Input {...field} value={field.value || ""} onKeyDown={handleArrowKey} className="bg-white" />
                    )}
                  />
                </div>
                <div className="space-y-1">
                  <Label className="text-xs uppercase text-gray-500">Ref. Std.</Label>
                  <Controller
                    name="ref_std"
                    control={form.control}
                    render={({ field }) => (
                      <Input {...field} value={field.value || ""} onKeyDown={handleArrowKey} className="bg-white" />
                    )}
                  />
                </div>
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
                        <td className="px-4 py-1 font-medium text-gray-900 border-r border-[#dee2e6] bg-gray-50/50">
                          {row.label}
                        </td>
                        <td className="p-0 border-r border-[#dee2e6]">
                          <Controller name={`core1.${row.key}` as any} control={form.control} render={({field}) => (
                            <Input {...field} value={field.value || ""} data-grid-row={idx} data-grid-col={0} onKeyDown={handleArrowKey} className="border-0 shadow-none h-8 rounded-none focus-visible:ring-1 focus-visible:ring-[#4a6fa5] bg-transparent px-3" />
                          )} />
                        </td>
                        <td className="p-0 border-r border-[#dee2e6]">
                          <Controller name={`core2.${row.key}` as any} control={form.control} render={({field}) => (
                            <Input {...field} value={field.value || ""} data-grid-row={idx} data-grid-col={1} onKeyDown={handleArrowKey} className="border-0 shadow-none h-8 rounded-none focus-visible:ring-1 focus-visible:ring-[#4a6fa5] bg-transparent px-3" />
                          )} />
                        </td>
                        <td className="p-0">
                          <Controller name={`core3.${row.key}` as any} control={form.control} render={({field}) => (
                            <Input {...field} value={field.value || ""} data-grid-row={idx} data-grid-col={2} onKeyDown={handleArrowKey} className="border-0 shadow-none h-8 rounded-none focus-visible:ring-1 focus-visible:ring-[#4a6fa5] bg-transparent px-3" />
                          )} />
                        </td>
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
                  { label: "GA Drg", key: "ga_drg" },
                  { label: "INS Class", key: "ins_class" },
                  { label: "Ref TI", key: "ref_ti" },
                  { label: "PRI Turns", key: "pri_turns" },
                  { label: "PRI Copper", key: "pri_copper" },
                  { label: "Former", key: "former" },
                  { label: "PRI Length", key: "pri_length" },
                  { label: "PRI Weight", key: "pri_weight" },
                  { label: "Sec. Terminal", key: "sec_terminal" },
                  { label: "Total Weight", key: "total_weight" }
                ].map((field) => (
                  <div key={field.key} className="space-y-1">
                    <Label className="text-xs uppercase text-gray-500">{field.label}</Label>
                    <Controller
                      name={field.key as any}
                      control={form.control}
                      render={({ field: f }) => (
                        <Input {...f} value={f.value || ""} data-field={field.dataField} onKeyDown={handleArrowKey} className="bg-white" />
                      )}
                    />
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
