# Gemini CT Design Software: Reverse-Engineering Analysis & Clone Plan

Source analysed: `Design soft/` (VB6 project sources, two compiled EXEs, Crystal
Reports, a sample report PDF and `Doc1.doc`, which turned out to hold 12 screenshots of the running app).

---

## 1. What is actually in `Design soft/`

| Item | What it is |
|---|---|
| `Code/*.frm, *.cls, *.bas` | VB6 source, about 31k lines. `CTfrm.frm` alone is 20,464 lines (5,392 of form layout and about 15k of code, 253 procedures). |
| `CT DESIGN 2012/GeminiTransformerDesign1.exe` | Project **GeminiDesigning**, 29 objects. **This build matches the source**: CTfrm has 253 methods, and PANEL_MASTERFRM and Panel_drg_Details are present. |
| `Code/VSS Design.exe` | An older build (project *VSSDesigning*, 27 objects, CTfrm with 251 methods). |
| `CT DESIGN 2012/Report*/ *.rpt` | Crystal Reports 8.x: `ctdesign5.rpt` (wound primary), `ctdesignwindow.rpt` (window/bar), `Costing.rpt`. |
| `Crystal Reports - ctdesign5.pdf` | A real printed design (GCT0010265). This is our **golden output sample**. |
| `Doc1.doc` | 12 screenshots: MDI shell, all 5 CTfrm tabs, the "Results..." form, the report preview, Mould, Core Size, Panel Master and Primary Winding forms. |

Both EXEs are **native-compiled VB6 (not P-code)**. I checked this by parsing the VB5! header (`lpNativeCode != 0`).

### 1.1 Critical gap: the source is incomplete

Comparing the EXE object table with the `.frm/.cls/.bas` files shows that these modules are compiled into
the EXE but **have no source in the folder**:

| Missing module | Kind | Methods | Why it matters |
|---|---|---|---|
| **`Individualfrm`** | Form ("Results...") | 24 | **The per-core secondary calculation.** It produces flux density, Ie, VK, RCT, ratio/phase error, composite error, ISF, SC temperature rise, and on "Accepted" it fills hidden grid `G1`, which is everything that gets saved and printed. |
| **`MeanBup`** | Module | 30 | Core geometry and current-density formulas: `prcArea`, `prcAreaEff`, `prcCalBU`, `prcCalMean`, `prcCalMeanCircular`, `prcCalSecC`, `prcCalNorSecC`, `prcCalRate`, `prcCalPC`, `prcCalNorPC`, `prcCalDElSTC`, `prcCalDElRate`, `prcAreaCSNormal`, `prcAreaCSSCur`, `prcDrgCheck`, and global variables (`ModDesign`, `DesignNo`, `RatioWatch()`, `CoreTap`, `DiaType`, `SetPTurnComm`...). |
| **`Report`** | Class | 15 | `ReportDesign(code)` builds the `designreport` / `design_reperror` tables that the Crystal report prints. |
| `Saving` | Class | 5 | Save helpers. |
| `ModifyDes` | Class | 1 | `ModifingDB(DesignNo)`: loads a saved design back into CTfrm (Modify Design). |
| `MDIDesign` | MDI form | 17 | Shell: menus, toolbar and status bar. Its panels are written to from everywhere. |
| `PrimaryCondfrm` | Form | 76 | Primary Winding Details master (screenshot available). Maintains `primarycond11`. |
| `Mouldfrm` | Form | 65 | Mould Details master (screenshot available). Maintains `mouldvol2`. |
| `searchfrm`, `C` (searchResultsfrm), `ModifyDesfrm` | Forms | 31 / 11 / 8 | Design search, results, modify picker. |
| `frmSplash` | Form | 3 | Splash. |

Other gaps:
* `CTfrm.frx` is from a different version (1.5 KB, while the `.frm` points to offsets up to 14 KB). As a result, the **static combo
  lists stored in the .frx** (Time, Freq, Core Shape, Core Material, VA, Ie@...) are not recoverable from source and
  must come from the EXE's form resources.
* **No database.** The app connects to ODBC DSN `ENHANCE` (SQL Server, `server.txt` = `server\server`). Every
  dropdown and every lookup table (conductors, losses/B-H curves, mould volumes, core sizes, HSV/IL/STC...) lives there.

**Without the missing modules plus the database, we cannot produce the same numbers as the EXE.** See §8.

---

## 2. Original application architecture

```
Sub Main (PragatiMain.bas)
 └─ prcConnect  → ADODB "Provider=MSDASQL.1;Data Source=ENHANCE"  (global con)
 └─ MDIDesign.Show   "GEMINI DESIGING SOFTWARE"
     Menus: Administrator | Design | Utility
       mnuAdmin, mnuDesign(mnuCtDesg, mnuNDesi, mnuModDes/mnuModDesi), mnuUtility,
       mnuSecCon, mnuComCore, mnuMould, mnuCosting, mnuPanelMst, mnuAbout, mnuExit
     Toolbar: C.T.Design | Print Design | Conductor ▾ | Core Size | Mould | Search | Calculator
     StatusBar: [1]=date [2]=field hint ("Set Class", "Select Core No"...) [3] [4]=CAPS [5]=CT kind ("Wound HTCT")
     MDI children:
       CTfrm               H.T. Current Transformer Designing (main workflow, 5 tabs)
         └ Individualfrm   "Results..." per-core calculation (modal-ish child)
         └ searchResultsfrm
       PrintDesfrm         Print Design (pick Des No → Crystal report)
       SecCondfrm          Secondary Conductor master (secconductor)
       PrimaryCondfrm      Primary Winding Details master (primarycond11)   [no source]
       CoreSizefrm         Common Core Dimension master (coresize)
       Mouldfrm            Mould Details master (mouldvol2)                 [no source]
       COSTMASTERfrm       Cost Master (new_costmaster, mst_part, itemmast)
       PANEL_MASTERFRM     Panel Master (panel_master)
       Panel_drg_Details   Panel ↔ drawing mapping
       STDesignSearchfrm / searchfrm / ModifyDesfrm  design lookup
       CommandDetailfrm    commandb
       frmAbout
```

---

## 3. The main workflow: CTfrm data flow

CTfrm is one MDI child with `SSTab1` (5 tabs). All state lives in **controls and hidden MSFlexGrids**. There is
no model layer. In the clone the state model has to mirror these controls exactly.

### Tab 1: C.T. Specifications [&1]
Inputs → `cmbEquipType` (from `Dtype where mode='2'`), `txtnoofcores` (1–5), `cmbNSV` (from `hsv`) → fills
`cmbHSV` (`hsv where nsv like`) and `CMBINSULATION` (`hsv.insclass, insclass1`) and calls `prcCoreSize` (Grid3 standard
cores), `cmbHSV` → `cmbIL` (`il where il like`), `cmbSTC` (`stc`), `cmbTime`, `cmbFreq`, `txtRateConti`.
Per-core ratios: `txtRatioDum(i)` (primary, e.g. `200-100`) / `txtRatioR2Dum(i)` (secondary). These are mirrored into
`txtRatio(i)` sorted ascending by `ChangeRa.prcInterChangeRatio` (through scratch table `ratiocheck`).

Per-core spec entry: `cmbnoofcores` (select core) → `cmbCType` (Metering / Protection / Special Protection) changes
which frame is shown and the class/ISF lists:
* Metering: Class 0.1, 0.2, 0.5, 1.0, 3.0, 5.0; ISF 5, 10, 15, 20, 25
* Protection: Class 5P, 10P, 15P; ALF 10, 15, 20, 30
* Special Protection (PS): VK ≥, RCT(75°C) ≤, Ie ≤ mA @ (Vk/2, Vk/4...)

Accuracy for multi-ratio cores: `OptComm` (common) or `OptInd` (different per tap, chosen through `IvAcc`).
`<< In` (`cmdin_Click`) appends a row to **`Grid`** (core spec grid, columns: 0 core no, 1 ratio, 2 equip, 3 hsv, 4 il,
5 stc, 6 time, 7 VA, 8 class, 9 ISF, 10 ALF, 11 VK, 12 RCT, 13 Ie, 14 Ie@, 16 "A"=individual accuracy, 18 tap no,
19 core type). `>> Del` removes a row, `Search` → searchResultsfrm, `Can.` → cancel design.

### Tab 2: Selection of Core(s) [&2]
`cmbStyle` = Rectangluar *(sic)* / Round. `OptCom(0/1)` = common or different core size per core (`cmbCore`).
Rectangular: OL × OW, WL × WW (or double-click a standard size in **Grid3** from `coresize`). Round: OD, ID.
`>>` (`cmdICore`) → **GCore** grid (core, OL, OW, WL, WW, OD, ID). `<<` removes.

### Tab 3: Primary Details [&3] (Wound types only; Bar/Window skip it, see `SSTab1_Click`)
`cmbConStyle` (MSForms multi-column combo: type Bare/Paper, style Single/Double/Double-Different/Double-Equal...) →
`PrcCalPrimaryDetails` queries `primarycond11` for valid turns (`txtTurns` list), width, C/S, side/top winding ht.
Computed: Min. Required C/S (`txtminCs`), Normal current density (`txtDelR`), STC current density (`txtDelStc`, limit 220),
I-dynamic kA = STC kA × 2.5 (`txtIDyn`), I-dynamic AT (`txtIDATurns`). `lstCond` = cross-section list, and choosing one fills
Total Conductors / Conductor per Coil (e.g. `10X0.8X(6)X2/10X0.8X(5)X2`).

### Tab 4: Secondary Details [&4] (`SSTab2`: Core Details | All Core Details)
Select core (`cmbSelctCore`) → `cmbSelctCore_Change` (676 lines) derives core geometry (Build-up, Area, Aeff, Mean
path, Weight) through the **MeanBup** formulas and proposes secondary turns, tap, conductor SWG (`secconductor`) and
compensation (`CrossSection()`). The user sets Height, Core Material (`cmbMetal`), Conductor (SWG + S/D/T/Q) and an optional
"different conductor per tap" (`ChK2` → `SWG` grid).
`Calculate` (`cmCal_Click`) → validates → `Individualfrm.Show` + `prcFillInd` (copies everything into the
Results form) → **Individualfrm computes**, the user presses **Accepted** → results are written into the
**hidden grid `G1` (110 rows × one column per core/tap)**. `G1` is the single source of truth for save and print.

### Tab 5: Dimensions [&5]
`prcDimen` / `prcWindDiemRound` / `prcWinddDiemRec` → **wounded** grid (secondary-wound OL×OW, window, height,
build-up, Max-Dim row). Model/Drg No (`txtdrg`, `drg`/`mouldvol2`) + Type (`cmbType`) → `Cal.` (`cmbCal_Click`)
→ **Former**, **Mould**, **Casting** grids, window clearances (`WOR1A/WOR2A`), primary winding mean length / length /
Cu weight / R20 / R75 / Watt loss (`WatLos`), weights: Core, Sec Cu, Pri Cu, Resin (`Weights.CalResinWt` uses
`mouldvol2.t_mould_vo`, Cu density 8.9, core 7.65, resin 1.8). Secondary insulation `CMBINSULATION`.
"Go Back To Stage" 1–4 buttons. **Save** (`cmdSave_Click`).

### Save → DB → Report
```
cmdSave_Click
  prcCheckingSave (validation)
  Dno = AutoDNo.AutoGen()        ← gen table, charecter='DEGNO', "CT" & Format(n,"0000000")
        (or ModifyDes + txtDesNo when ModDesign="MOD")
  INSERT designspec   one row per Grid row   (spec + core sizes + primary + former + casting + weights + final dims)
  INSERT designerror  one row per G1 column   (ratio/phase/composite error tables @ 120/100/20/5 %, full & quarter VA)
  INSERT designcore   one row per G1 column   (sec. turns, layers, wind ht, RCT75, flux, VK, Ie, CE, ISF, SC dens...)
  → "Preview?" → Report.ReportDesign(code)  [no source] → designreport/design_reperror
  → CrystalReport1 (ctdesign5.rpt | ctdesignwindow.rpt)
```

### `G1` row map (from `cmdSave_Click`; this becomes the TS `CoreResult` type)
| Row | Field | Row | Field |
|---|---|---|---|
| 0 | core_ratio | 1 | core no |
| 2 | tap_ratio | 3 | r_turns |
| 4 / 5 | pri / sec current | 6 / 7 | SWG no (S/D/T/Q) / SWG |
| 8 | layers | 9 | winding ht |
| 10 | RCT75 | 11, 21 | excitation current |
| 12 | VK | 13 / 14 | C1 / C2 |
| 15 / 16 | Ie nor / Ie ALF | 17 | ISF |
| 18 / 19 / 20 | SC curr dens / sat dens / SC temp rise | 22 / 23 | CE nor / CE ALF |
| 24 | flux | 25 | sp_alf |
| 26 / 27 / 28 | BUP / Aeff (area) / Mean | 29, 32 | ht |
| 30 / 31 / 33 / 34 / 35 | sp rct / ie / va / class / isf | 36–47 | ratio, composite and phase errors (full VA) |
| 48 / 49 | core type / style | 51–56 | OD, ID, OL, OW, WL, WW |
| 57–68 | errors (quarter VA) | 69 / 70 | name1 / name2 |
| 71 / 72 | core ht / total layers | 73–79 | wound OD, ID, OL, OW, WL, WW, Ht |
| 80 / 81 | sec Cu wt / core wt | 82–91 | %Ipn points (full / quarter) |
| 92–101 | Bm at each point | 102 / 103 / 104 | sec cond / dia / min req C/S |
| 105 / 106 / 107 | between turns / compensation / core material | 108 / 109 | tap / max tap |

---

## 4. Scenario matrix (what "multiple scenarios" means in the code)

| Axis | Values | Where it branches |
|---|---|---|
| Equipment type | Wound (Wo-C / Wo-R), Window (Win-C / Win-R), Bar (B-C / B-R) | `cmbEquipType_*`, `SSTab1_Click` (skip tab 3), `CalResinWt`, save (casting rows differ), report choice |
| Core shape | Rectangular / Round | `cmbStyle`, `CalLayers*` vs `CalLayers*Round`, `prcWinddDiemRec` vs `prcWindDiemRound` |
| Core type | Metering / Protection / Special Protection (PS) | spec frames, Individualfrm visibility, BM/ISF vs composite error vs C1/C2/VK |
| Ratio | single, or multi-ratio taps (`200-100`) | `InStr(ratio,"-")`, `CalLayersTapS/TapDiff[Round]`, turn markings `1S1-1S2...` |
| Accuracy | common or different per tap | `OptComm` / `OptInd`, Grid col 16 = "A" |
| Secondary conductor | same for all taps / different per tap (`ChK2`) | `DiaType` "S"/"D", `SWG` grid |
| Conductor parallel | S / D / T / Q strands | `txtNo`, `CrossSection()` compensation |
| Primary conductor style | Bare Single/Double, Paper Single/Double-Different/Double-Equal, Series, BAR-C/R | `PrcCalPrimaryDetails`, `SearchForStyle` rules (1 turn vs >1 turns) |
| Frequency | 50 / 60 Hz | BM factor 4.5 vs 3.75, C1 divisor 3000 vs 2500 |
| Rated continuous current | given / blank | `prcCalRate` vs `prcCalNorSecC`, `WatLos` |
| Core material | CRGO-B, MU metal... | `losess` / `MUMAGCUR` B-H lookups, density |
| New vs Modify design | `ModDesign="MOD"` | `Form_Load` → `ModifyDes.ModifingDB`, `ModCal` |

---

## 5. Formula catalogue (recovered so far, with source refs)

| Formula | Source |
|---|---|
| BM@50 = (VA + I²·R75)·4.5·10⁵ / (AT·Aeff); @60 uses 3.75 | `comboPrc.CalProtectingHt` |
| R20 = 0.0177·(L/1000)/A; R75 = R20·1.2161 | `comboPrc` |
| Winding ht = 1.5 + OD·layers + (layers−1)·0.25 + (1.5+3) | `comboPrc`, `ChK2_Click` |
| MLT = Wht·3.142 + 2·BUP + 2·Ht | `comboPrc` |
| ISF = 21000/BM; C1 = Aeff·N/3000 (50 Hz) or /2500; C2 = 1.3·Mean/N; VK = C1·96; Ie @Vk/4 = 0.092·C2, @Vk/2 = 0.145·C2, else 0.34·C2 (×1000 mA) | `comboPrc` (commented copy; live version is in Individualfrm) |
| Burden angle: cosλ = 0.8 if VA>5 else 1, tanδ = VA·sinλ/(I²R75 + VA·cosλ) | `CTfrm.Trig` |
| Compensation % = 100/(turns·Np/Is) and the S/D/T/Q split ±20 % search | `CTfrm.CrossSection` |
| I-dyn = STC kA × 2.5 | `cmbConStyle_Click` |
| Pri watt loss = Imax²·R | `CTfrm.WatLos` |
| Resin wt = (Vmould − Wcore/7.65·1000 − Wsec/8.9·1000)·1.8/1000 | `Weights.CalResinWt` |
| Resin thickness per HSV | `prcResinThick` (`hsv.resin`) |
| Design no. = "CT"+7 digits (gen table) | `AutoDNo` |
| Number to words (Indian lakh/crore) | `PragatiMain.convert` |
| Costing (core by shape/weight, Cu, resin, coating) | `costing.cls` (4,795 lines) |

There are also **known source bugs that affect output**, e.g. `WatLos` compares against the undefined `MaxButton`
(so it takes the *last* ratio rather than the max), and `Mid(s, i, i)` is used where `Mid(s, i, 1)` was meant. An exact clone
reproduces these. They are listed in a quirks register (§7.3).

---

## 6. Database (inferred from ~200 SQL strings in source + EXE)

**Masters / lookups:** `Dtype`(item, mode), `hsv`(nsv, hsv, insclass, insclass1, resin), `il`, `stc`, `freq`,
`coresize`(nsv, OL, OW, WL, WW), `secconductor`(t_swg, t_cs, t_ods, ...), `primarycond11`(prid, type, style, mod, drg,
t_np, t_csp, t_cond_siz, t_cond_coi, t_whps, t_whpt, pri_wind_width), `losess`(t_bm, t_aw0...), `MUMAGCUR`(t_bm, t_awm,
t_aww), `normalCD`(ncd), `mouldvol2`(t_drg, t_DrgType, t_mould_vo, t_moulddrg...), `drg`, `Drgtypes`, `commanDB`,
`mst_part`, `itemmast`, `costmaster`, `new_costmaster`, `core_rate`, `panel_master`, `panel_drg_details`.
**Transactional:** `designspec`, `designcore`, `designerror`, `designreport`, `design_reperror`, `costing`,
`costingcomm`, `gen` (counters), `ratiocheck` (scratch).

The full column list per table will be generated mechanically from the `rst!field` usages (next deliverable).

---

## 7. Target architecture for the clone

### 7.1 Stack
* **Frontend:** this repo (React 18 + Vite + TS + Tailwind/shadcn). We add a "classic" theme that reproduces the
  Win2000/XP look (MDI window chrome, SSTab tabs, MSFlexGrid, 3D frames, Verdana 8.25) at pixel-equivalent layout.
  VB twips → px (÷15) are taken directly from the `.frm` `Left/Top/Width/Height`.
* **Backend (recommended):** a small Node/Express (or .NET) API over **the same SQL Server schema**, so the clone can run
  side-by-side with the old EXE against one DB and existing designs and masters keep working. A browser-only
  (localStorage) mode is possible for demos but can't share designs or masters.
* **Reports:** PDF built to the Crystal layout (`ctdesign5`, `ctdesignwindow`, `Costing`), verified against
  `Crystal Reports - ctdesign5.pdf`.

### 7.2 Code layout
```
src/design/
  vb/            VB6 runtime semantics: Val, CDbl, CInt, Round (banker's), Format, Mid/InStr/Left, string↔number
                 coercion, Variant Empty, "Mid(s,i,i)" behaviour …   ← required for bit-exact numbers
  controls/      VB-like widgets: VbForm, SSTab, Frame, FlexGrid(TextMatrix,Rows,Cols,Tag), VbCombo, VbText, OptionButton
  state/         CTfrm state = one object holding every control value + hidden grids (Grid, GCore, G1, G2,
                 SWG, GCross, wounded, Former, Mould, Casting, Grid3, TestGrid) — mirrors the form 1:1
  engine/        1:1 ports, same procedure names:
                   ctfrm/*.ts      (cmbSelctCore_Change, CalLayersTapS, prcDimen, cmbCal_Click, cmdSave_Click …)
                   meanbup.ts, individual.ts, report.ts   (from EXE recovery)
                   comboPrc.ts, weights.ts, changeRa.ts, autoDNo.ts, costing.ts, pragatiMain.ts
  forms/         React screens: MDIDesign, CTfrm (5 tabs), Individualfrm, PrintDesfrm, SecCondfrm,
                 PrimaryCondfrm, CoreSizefrm, Mouldfrm, COSTMASTERfrm, PANEL_MASTERFRM, Panel_drg_Details, search…
  db/            repository layer: one function per SQL statement in the VB code
server/          API + SQL (same table/column names)
```
Event handlers stay event handlers: `cmbNSV_Change` → `onNSVChange(state)`. This keeps the side-effect order,
which matters a lot here, because Change/Click/GotFocus/LostFocus chains do the actual computing.

### 7.3 Fidelity rules ("exact clone")
1. Every VB procedure gets a TS function with the **same name**, and the VB file and line range go in a header comment.
2. VB semantics come from the `vb/` helpers. No "cleaner" JS maths: VB `Round` is banker's rounding, and `CDbl("")` throws.
3. Bugs and quirks are reproduced by default and listed in `QUIRKS.md` with a toggle, so fixes are opt-in.
4. Golden tests: the same inputs fed to the old EXE and to the clone must give identical `designspec/designcore/designerror` rows
   and the same report. The first golden case is design GCT0010265 from the sample PDF.
5. Every MsgBox text, status-bar hint, caption, tab order and enable/disable rule is kept.

---

## 8. Recovering what isn't in the source

Option A is strongly preferred because it is the only route to "every single line":
* **A. Get the complete VB6 project** (the `.vbp`, `Individualfrm.frm/.frx`, `MeanBup.bas`, `Report.cls`, `Saving.cls`,
  `ModifyDes.cls`, `MDIDesign.frm`, `PrimaryCondfrm.frm`, `Mouldfrm.frm`, `searchfrm.frm`, `searchResultsfrm.frm`,
  `ModifyDesfrm.frm`, `frmSplash.frm`, and matching `.frx` files). It is usually on the original developer's PC or in VSS.
* **B. Recover from the EXE.** Form layouts, control trees and static lists can be extracted exactly. Logic has to be
  decompiled from native x86 (Ghidra / VB Decompiler). That gives correct formulas but not the original lines, and
  it needs heavy verification against golden runs of the EXE.
* **Database:** we need a backup (`.bak`) or at least an export of the master tables (§6), plus a few
  saved designs for golden tests. Without it, dropdowns are empty and lookups (conductors, B-H losses, mould volumes)
  can't be done.

---

## 9. Phased plan

| Phase | Deliverable |
|---|---|
| 0 | Get missing source + DB (or start EXE recovery). Set up golden cases on the old EXE. |
| 1 | Full inventory: control tree for every form (twips→px JSON), every SQL statement → repository fn, every table/column, G1 map, static combo lists. |
| 2 | `vb/` runtime + `controls/` widgets + MDI shell (menus, toolbar, status bar). |
| 3 | Masters: SecCondfrm, CoreSizefrm, PrimaryCondfrm, Mouldfrm, Panel, Cost master (simple CRUD, which also proves the DB layer). |
| 4 | CTfrm tabs 1–2 (spec, core grid, core selection). |
| 5 | CTfrm tab 3 (primary) + tab 4 + Individualfrm (the calculation core): line-by-line port with golden tests per scenario in §4. |
| 6 | Tab 5 dimensions/weights, Save (AutoDNo, 3 tables), Modify design, search. |
| 7 | Report class + PDF reports (wound / window / costing), Print Design. |
| 8 | Costing (`costing.cls`), panel modules, side-by-side acceptance against the EXE. |
