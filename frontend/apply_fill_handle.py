import re

filepath = 'src/app/pages/sheet-editor/sheet-editor.component.ts'
content = open(filepath, encoding='utf-8').read()

# ── 1. ADD CSS for fill popup (after existing fill-handle style) ──────────
old_css = "    .fill-handle { background:#34a853; border:2px solid #fff; border-radius:50%; bottom:-5px; right:-5px; cursor:crosshair; height:8px; position:absolute; width:8px; z-index:30; box-shadow:0 1px 3px rgba(0,0,0,.4); }"
new_css = old_css + """
    .fill-handle:hover { transform:scale(1.4); transition:transform .1s; }
    .cell.fill-preview { outline: 1px dashed #34a853 !important; outline-offset: -1px; background-color: rgba(52,168,83,0.08) !important; }
    .fill-options-wrap { position:fixed; z-index:200000; pointer-events:none; }
    .fill-options-icon { width:22px; height:22px; background:#fff; border:1px solid #bbb; border-radius:3px; display:flex; align-items:center; justify-content:center; cursor:pointer; pointer-events:all; box-shadow:0 1px 4px rgba(0,0,0,.18); }
    .fill-options-icon:hover { background:#f1f3f4; }
    .fill-options-icon svg { width:14px; height:14px; }
    .fill-options-dropdown { position:absolute; top:26px; left:0; background:#202124; border-radius:8px; box-shadow:0 6px 24px rgba(0,0,0,.35); min-width:210px; padding:6px 0; pointer-events:all; }
    .fill-opt-item { display:flex; align-items:center; gap:8px; padding:9px 16px; font-size:13px; color:#e8eaed; cursor:pointer; white-space:nowrap; }
    .fill-opt-item:hover { background:rgba(255,255,255,.1); }
    .fill-opt-check { width:16px; font-size:14px; color:#34a853; flex-shrink:0; }"""
content = content.replace(old_css, new_css)

# ── 2. ADD fill popup HTML after grid div closes (after </div> at grid end) ──
popup_html = """
      <!-- ═══ FILL OPTIONS POPUP ════════════════════════════════════════════ -->
      <div class="fill-options-wrap" *ngIf="fillPopupState"
           [style.left.px]="fillPopupState.iconX"
           [style.top.px]="fillPopupState.iconY">
        <div class="fill-options-icon" (click)="fillPopupState.showMenu = !fillPopupState.showMenu; $event.stopPropagation()">
          <svg viewBox="0 0 16 16" fill="none" stroke="#5f6368" stroke-width="1.5"><rect x="2" y="2" width="5" height="5" rx="1"/><rect x="9" y="2" width="5" height="5" rx="1"/><rect x="2" y="9" width="5" height="5" rx="1"/><rect x="9" y="9" width="5" height="5" rx="1"/></svg>
        </div>
        <div class="fill-options-dropdown" *ngIf="fillPopupState.showMenu">
          <div class="fill-opt-item" (click)="executeFillMode('Fill Series')">
            <span class="fill-opt-check">{{ fillPopupState.mode === 'Fill Series' ? '✓' : '' }}</span>Fill Series
          </div>
          <div class="fill-opt-item" (click)="executeFillMode('Copy Cells')">
            <span class="fill-opt-check">{{ fillPopupState.mode === 'Copy Cells' ? '✓' : '' }}</span>Copy Cells
          </div>
          <div class="fill-opt-item" (click)="executeFillMode('Fill Formatting')">
            <span class="fill-opt-check">{{ fillPopupState.mode === 'Fill Formatting' ? '✓' : '' }}</span>Fill Formatting
          </div>
          <div class="fill-opt-item" (click)="executeFillMode('Fill Without Formatting')">
            <span class="fill-opt-check">{{ fillPopupState.mode === 'Fill Without Formatting' ? '✓' : '' }}</span>Fill Without Formatting
          </div>
        </div>
      </div>
"""
# Insert after the grid scroll div closes
content = content.replace("      <!-- Right Side Panel for Apps -->", popup_html + "      <!-- Right Side Panel for Apps -->")

# ── 3. ADD state properties after 'private isFilling = false;' ───────────
old_state = "  private isFilling = false;\n  fillEnd: { r: number, c: number } | null = null;\n  private fillStart: { r: number, c: number } | null = null;"
new_state = """  private isFilling = false;
  fillEnd: { r: number, c: number } | null = null;
  private fillStart: { r: number, c: number } | null = null;
  fillPopupState: {
    srcMinR: number; srcMaxR: number; srcMinC: number; srcMaxC: number;
    dstMinR: number; dstMaxR: number; dstMinC: number; dstMaxC: number;
    goDown: boolean; goUp: boolean; goRight: boolean; goLeft: boolean;
    ctrlKey: boolean;
    sourceData: { r: number; c: number; val: string; fmt: any }[];
    targetBackup: { r: number; c: number; val: string; fmt: any }[];
    mode: 'Fill Series' | 'Copy Cells' | 'Fill Formatting' | 'Fill Without Formatting';
    showMenu: boolean; iconX: number; iconY: number;
  } | null = null;"""
content = content.replace(old_state, new_state)

# ── 4. UPDATE onDocMouseUp to pass ctrlKey and show popup ────────────────
old_mouseup = """  onDocMouseUp() {
    if (this.isFilling && this.fillEnd) {
      this.applyFill();
    }
    this.isDraggingRange = false;
    this.isFilling = false;
    this.fillStart = null;
  }"""
new_mouseup = """  onDocMouseUp(e?: MouseEvent) {
    if (this.isFilling && this.fillEnd) {
      this.applyFill(e?.ctrlKey ?? false);
    }
    this.isDraggingRange = false;
    this.isFilling = false;
    this.fillStart = null;
  }"""
content = content.replace(old_mouseup, new_mouseup)

# ── 5. UPDATE onDocClick to close fill popup ────────────────────────────
old_docclick = "  onDocClick() { this.closeMenus(); this.activePalette = null; this.hideCtx(); }"
new_docclick = "  onDocClick() { this.closeMenus(); this.activePalette = null; this.hideCtx(); if (this.fillPopupState) this.fillPopupState.showMenu = false; }"
content = content.replace(old_docclick, new_docclick)

# ── 6. ALSO clear fillPopup on cell mousedown ────────────────────────────
old_celldown = "  onCellMouseDown(e: MouseEvent, r: number, c: number) {\n    if ((e.target as HTMLElement).classList.contains('fill-handle')) return;"
new_celldown = """  onCellMouseDown(e: MouseEvent, r: number, c: number) {
    if (this.fillPopupState) this.fillPopupState = null;
    if ((e.target as HTMLElement).classList.contains('fill-handle')) return;"""
content = content.replace(old_celldown, new_celldown)

# ── 7. REPLACE applyFill with full Zoho-style implementation ─────────────
old_apply = """  private applyFill() {
    if (!this.fillStart || !this.fillEnd) return;
    // Determine the source cells (selection range or just one cell)
    const srcMinR = this.rangeStart ? Math.min(this.rangeStart.r, this.rangeEnd!.r) : this.fillStart.r;
    const srcMaxR = this.rangeStart ? Math.max(this.rangeStart.r, this.rangeEnd!.r) : this.fillStart.r;
    const srcMinC = this.rangeStart ? Math.min(this.rangeStart.c, this.rangeEnd!.c) : this.fillStart.c;
    const srcMaxC = this.rangeStart ? Math.max(this.rangeStart.c, this.rangeEnd!.c) : this.fillStart.c;

    const dstR = this.fillEnd.r;
    const dstC = this.fillEnd.c;

    this.pushHistory();

    // Determine fill direction
    const goDown = dstR > srcMaxR;
    const goUp = dstR < srcMinR;
    const goRight = dstC > srcMaxC;
    const goLeft = dstC < srcMinC;

    if (goDown || goUp) {
      // Fill column-wise
      for (let c = srcMinC; c <= srcMaxC; c++) {
        const srcVals: string[] = [];
        for (let r = srcMinR; r <= srcMaxR; r++) srcVals.push(this.cells[r][c]);
        if (goDown) {
          for (let r = srcMaxR + 1; r <= dstR; r++) {
            this.cells[r][c] = this.getNextSeriesValue(srcVals, r - srcMinR, true);
          }
        } else {
          for (let r = srcMinR - 1; r >= dstR; r--) {
            this.cells[r][c] = this.getNextSeriesValue(srcVals, r - srcMaxR, true);
          }
        }
      }
    } else if (goRight || goLeft) {
      // Fill row-wise
      for (let r = srcMinR; r <= srcMaxR; r++) {
        const srcVals: string[] = [];
        for (let c = srcMinC; c <= srcMaxC; c++) srcVals.push(this.cells[r][c]);
        if (goRight) {
          for (let c = srcMaxC + 1; c <= dstC; c++) {
            this.cells[r][c] = this.getNextSeriesValue(srcVals, c - srcMinC, false);
          }
        } else {
          for (let c = srcMinC - 1; c >= dstC; c--) {
            this.cells[r][c] = this.getNextSeriesValue(srcVals, c - srcMaxC, false);
          }
        }
      }
    }

    // Update selection to show the full filled area
    const newMinR = Math.min(srcMinR, dstR);
    const newMaxR = Math.max(srcMaxR, dstR);
    const newMinC = Math.min(srcMinC, dstC);
    const newMaxC = Math.max(srcMaxC, dstC);
    this.rangeStart = { r: newMinR, c: newMinC };
    this.rangeEnd = { r: newMaxR, c: newMaxC };
    this.selectedRow = newMinR;
    this.selectedCol = newMinC;
    this.fillEnd = null;
    this.onCellChange();
    this.save();
    this.showToast('Series filled.');
  }"""

new_apply = """  private applyFill(ctrlKey: boolean = false) {
    if (!this.fillStart || !this.fillEnd) return;

    const srcMinR = this.rangeStart ? Math.min(this.rangeStart.r, this.rangeEnd!.r) : this.fillStart.r;
    const srcMaxR = this.rangeStart ? Math.max(this.rangeStart.r, this.rangeEnd!.r) : this.fillStart.r;
    const srcMinC = this.rangeStart ? Math.min(this.rangeStart.c, this.rangeEnd!.c) : this.fillStart.c;
    const srcMaxC = this.rangeStart ? Math.max(this.rangeStart.c, this.rangeEnd!.c) : this.fillStart.c;

    const dstR = this.fillEnd.r;
    const dstC = this.fillEnd.c;

    const goDown  = dstR > srcMaxR;
    const goUp    = dstR < srcMinR;
    const goRight = dstC > srcMaxC;
    const goLeft  = dstC < srcMinC;

    if (!goDown && !goUp && !goRight && !goLeft) return;

    // snapshot source data + formats
    const sourceData: { r: number; c: number; val: string; fmt: any }[] = [];
    for (let r = srcMinR; r <= srcMaxR; r++)
      for (let c = srcMinC; c <= srcMaxC; c++)
        sourceData.push({ r, c, val: this.cells[r][c], fmt: JSON.parse(JSON.stringify(this.formats[`${r},${c}`] ?? null)) });

    // compute fill target bounds
    const dstMinR = goUp   ? dstR    : (goDown  ? srcMaxR + 1 : srcMinR);
    const dstMaxR = goUp   ? srcMinR - 1 : (goDown  ? dstR    : srcMaxR);
    const dstMinC = goLeft ? dstC    : (goRight ? srcMaxC + 1 : srcMinC);
    const dstMaxC = goLeft ? srcMinC - 1 : (goRight ? dstC    : srcMaxC);

    // backup target cells
    const targetBackup: { r: number; c: number; val: string; fmt: any }[] = [];
    for (let r = dstMinR; r <= dstMaxR; r++)
      for (let c = dstMinC; c <= dstMaxC; c++)
        targetBackup.push({ r, c, val: this.cells[r][c], fmt: JSON.parse(JSON.stringify(this.formats[`${r},${c}`] ?? null)) });

    this.pushHistory();

    // default mode: single number copies, multi-value series (Zoho default)
    const isSingleSrc = srcMinR === srcMaxR && srcMinC === srcMaxC;
    const singleVal = isSingleSrc ? (this.cells[srcMinR][srcMinC] ?? '') : '';
    const singleIsNum = isSingleSrc && !isNaN(Number(singleVal)) && singleVal.trim() !== '';
    let defaultMode: 'Fill Series' | 'Copy Cells' =
      (isSingleSrc && !singleIsNum && !singleVal.startsWith('=')) ? 'Copy Cells' :
      (!isSingleSrc) ? 'Fill Series' :
      ctrlKey ? 'Fill Series' : 'Copy Cells';
    // ctrl inverts: multi-src with ctrl → copy, single num with ctrl → series
    if (ctrlKey && isSingleSrc && singleIsNum) defaultMode = 'Fill Series';
    if (ctrlKey && !isSingleSrc) defaultMode = 'Copy Cells';

    this._doFill(srcMinR, srcMaxR, srcMinC, srcMaxC, dstMinR, dstMaxR, dstMinC, dstMaxC,
                 goDown, goUp, goRight, goLeft, defaultMode, ctrlKey);

    // update selection
    const newMinR = Math.min(srcMinR, dstMinR);
    const newMaxR = Math.max(srcMaxR, dstMaxR);
    const newMinC = Math.min(srcMinC, dstMinC);
    const newMaxC = Math.max(srcMaxC, dstMaxC);
    this.rangeStart = { r: newMinR, c: newMinC };
    this.rangeEnd   = { r: newMaxR, c: newMaxC };
    this.selectedRow = newMinR; this.selectedCol = newMinC;
    this.fillEnd = null;
    this.onCellChange();
    this.save();

    // compute popup icon position from last dst cell DOM
    const anchorEl = document.getElementById(`cell-${dstMaxR}-${dstMaxC}`);
    let iconX = 0, iconY = 0;
    if (anchorEl) {
      const rect = anchorEl.getBoundingClientRect();
      iconX = rect.right + 4;
      iconY = rect.bottom + 4;
    }

    this.fillPopupState = {
      srcMinR, srcMaxR, srcMinC, srcMaxC,
      dstMinR, dstMaxR, dstMinC, dstMaxC,
      goDown, goUp, goRight, goLeft, ctrlKey,
      sourceData, targetBackup,
      mode: defaultMode,
      showMenu: false, iconX, iconY
    };
  }

  private _doFill(
    srcMinR: number, srcMaxR: number, srcMinC: number, srcMaxC: number,
    dstMinR: number, dstMaxR: number, dstMinC: number, dstMaxC: number,
    goDown: boolean, goUp: boolean, goRight: boolean, goLeft: boolean,
    mode: 'Fill Series' | 'Copy Cells' | 'Fill Formatting' | 'Fill Without Formatting',
    ctrlKey: boolean
  ) {
    const getSrc = (r: number, c: number) =>
      this.fillPopupState?.sourceData.find(d => d.r === r && d.c === c) ??
      { r, c, val: this.cells[r][c], fmt: this.formats[`${r},${c}`] ?? null };

    const srcRows = srcMaxR - srcMinR + 1;
    const srcCols = srcMaxC - srcMinC + 1;

    if (goDown || goUp) {
      for (let c = srcMinC; c <= srcMaxC; c++) {
        const srcVals: string[] = [];
        for (let r = srcMinR; r <= srcMaxR; r++) srcVals.push(getSrc(r, c).val);
        const fillTargetRows = (startR: number, endR: number, step: number, reversed: boolean) => {
          let offset = 1;
          for (let r = startR; r !== endR + step; r += step, offset++) {
            const srcR = srcMinR + ((offset - 1) % srcRows);
            const srcItem = getSrc(srcR, c);
            if (mode === 'Fill Formatting') {
              if (srcItem.fmt) this.formats[`${r},${c}`] = JSON.parse(JSON.stringify(srcItem.fmt));
            } else if (mode === 'Copy Cells' || mode === 'Fill Without Formatting') {
              this.cells[r][c] = this._shiftFormula(srcItem.val, r - srcR, 0);
              if (mode === 'Copy Cells' && srcItem.fmt) this.formats[`${r},${c}`] = JSON.parse(JSON.stringify(srcItem.fmt));
            } else {
              // Fill Series
              const relOffset = reversed ? (r - (srcMinR - 1)) : (r - srcMaxR);
              this.cells[r][c] = this._getNextSeriesEx(srcVals, relOffset, true, ctrlKey);
              if (srcItem.fmt) this.formats[`${r},${c}`] = JSON.parse(JSON.stringify(srcItem.fmt));
            }
          }
        };
        if (goDown) fillTargetRows(dstMinR, dstMaxR,  1, false);
        else        fillTargetRows(dstMaxR, dstMinR, -1, true);
      }
    } else if (goRight || goLeft) {
      for (let r = srcMinR; r <= srcMaxR; r++) {
        const srcVals: string[] = [];
        for (let c = srcMinC; c <= srcMaxC; c++) srcVals.push(getSrc(r, c).val);
        const fillTargetCols = (startC: number, endC: number, step: number, reversed: boolean) => {
          let offset = 1;
          for (let c = startC; c !== endC + step; c += step, offset++) {
            const srcC = srcMinC + ((offset - 1) % srcCols);
            const srcItem = getSrc(r, srcC);
            if (mode === 'Fill Formatting') {
              if (srcItem.fmt) this.formats[`${r},${c}`] = JSON.parse(JSON.stringify(srcItem.fmt));
            } else if (mode === 'Copy Cells' || mode === 'Fill Without Formatting') {
              this.cells[r][c] = this._shiftFormula(srcItem.val, 0, c - srcC);
              if (mode === 'Copy Cells' && srcItem.fmt) this.formats[`${r},${c}`] = JSON.parse(JSON.stringify(srcItem.fmt));
            } else {
              const relOffset = reversed ? (c - (srcMinC - 1)) : (c - srcMaxC);
              this.cells[r][c] = this._getNextSeriesEx(srcVals, relOffset, false, ctrlKey);
              if (srcItem.fmt) this.formats[`${r},${c}`] = JSON.parse(JSON.stringify(srcItem.fmt));
            }
          }
        };
        if (goRight) fillTargetCols(dstMinC, dstMaxC,  1, false);
        else         fillTargetCols(dstMaxC, dstMinC, -1, true);
      }
    }
  }

  executeFillMode(mode: 'Fill Series' | 'Copy Cells' | 'Fill Formatting' | 'Fill Without Formatting') {
    if (!this.fillPopupState) return;
    const p = this.fillPopupState;

    // restore target cells from backup
    for (const b of p.targetBackup) {
      this.cells[b.r][b.c] = b.val;
      if (b.fmt) this.formats[`${b.r},${b.c}`] = JSON.parse(JSON.stringify(b.fmt));
      else delete this.formats[`${b.r},${b.c}`];
    }
    // restore source formatting if needed
    for (const s of p.sourceData) {
      if (s.fmt) this.formats[`${s.r},${s.c}`] = JSON.parse(JSON.stringify(s.fmt));
      else delete this.formats[`${s.r},${s.c}`];
      this.cells[s.r][s.c] = s.val;
    }

    this._doFill(p.srcMinR, p.srcMaxR, p.srcMinC, p.srcMaxC,
                 p.dstMinR, p.dstMaxR, p.dstMinC, p.dstMaxC,
                 p.goDown, p.goUp, p.goRight, p.goLeft, mode, p.ctrlKey);

    this.fillPopupState = { ...p, mode, showMenu: false };
    this.onCellChange();
    this.save();
  }

  private _shiftFormula(val: string, rowDelta: number, colDelta: number): string {
    if (!val || !val.startsWith('=')) return val;
    return val.replace(/\\$?[A-Z]+\\$?\\d+/g, (match) => {
      if (match.startsWith('$') && match.includes('$', 1)) return match; // absolute
      const hasAbsCol = match.startsWith('$');
      const hasAbsRow = match.match(/\\$\\d+$/);
      const colStr = match.replace(/\\$/g, '').match(/^[A-Z]+/)![0];
      const rowStr = match.match(/\\d+$/)![0];
      let colIdx = 0;
      for (let i = 0; i < colStr.length; i++) colIdx = colIdx * 26 + (colStr.charCodeAt(i) - 64);
      colIdx -= 1;
      let rowIdx = parseInt(rowStr, 10) - 1;
      if (!hasAbsCol) colIdx += colDelta;
      if (!hasAbsRow) rowIdx += rowDelta;
      if (colIdx < 0) colIdx = 0;
      if (rowIdx < 0) rowIdx = 0;
      return `${colName(colIdx)}${rowIdx + 1}`;
    });
  }

  private _getNextSeriesEx(srcVals: string[], offsetSteps: number, isVertical: boolean, ctrlKey: boolean): string {
    const srcLen = srcVals.length;
    const isSingle = srcLen === 1;
    const isNum = (s: string) => s.trim() !== '' && !isNaN(Number(s));

    // Ctrl inverts: single copy→series, multi series→copy
    const trySeries = isSingle ? ctrlKey : !ctrlKey;

    const cycleIdx = ((offsetSteps - 1) % srcLen + srcLen) % srcLen;
    let v = srcVals[cycleIdx];

    if (v.startsWith('=')) {
      const shiftAmt = offsetSteps - 1 - cycleIdx;
      const rd = isVertical ? shiftAmt : 0;
      const cd = isVertical ? 0 : shiftAmt;
      return this._shiftFormula(v, rd, cd);
    }

    if (!trySeries) return v; // Copy mode

    // Number series
    if (isSingle && isNum(v)) {
      return String(Number(v) + offsetSteps - 1);
    }
    if (!isSingle) {
      const nums = srcVals.map(s => Number(s));
      if (nums.every((n, i) => isNum(srcVals[i]))) {
        const step = nums.length > 1 ? nums[1] - nums[0] : 1;
        return String(nums[nums.length - 1] + step * (offsetSteps - srcLen + srcLen - srcLen + (offsetSteps - srcLen)));
      }
    }

    // Multi-value number series (robust)
    if (!isSingle) {
      const nums = srcVals.map(Number);
      if (nums.every((n, i) => isNum(srcVals[i]))) {
        const step = nums.length > 1 ? nums[1] - nums[0] : 1;
        const target = nums[0] + step * (offsetSteps - 1 + (srcLen - 1));
        return String(target);
      }
    }

    // Date series MM-DD-YY or MM-DD-YYYY
    const dateRx = /^(\\d{1,2})-(\\d{1,2})-(\\d{2,4})$/;
    const parseDate = (s: string) => {
      const m = s.match(dateRx);
      if (!m) return null;
      const mon = +m[1], d = +m[2];
      let y = +m[3]; if (y < 100) y += 2000;
      return { y, m: mon, d, twoDigit: m[3].length === 2 };
    };
    if (isSingle) {
      const pd = parseDate(v);
      if (pd) {
        const dt = new Date(pd.y, pd.m - 1, pd.d);
        dt.setDate(dt.getDate() + offsetSteps - 1);
        const nm = String(dt.getMonth() + 1).padStart(2, '0');
        const nd = String(dt.getDate()).padStart(2, '0');
        const ny = pd.twoDigit ? String(dt.getFullYear()).slice(2) : String(dt.getFullYear());
        return `${nm}-${nd}-${ny}`;
      }
    } else {
      const d0 = parseDate(srcVals[0]), d1 = parseDate(srcVals[1]);
      if (d0 && d1) {
        const dt0 = new Date(d0.y, d0.m - 1, d0.d);
        const dt1 = new Date(d1.y, d1.m - 1, d1.d);
        const diffDays = Math.round((dt1.getTime() - dt0.getTime()) / 86400000);
        const dt = new Date(d0.y, d0.m - 1, d0.d);
        dt.setDate(dt.getDate() + diffDays * (offsetSteps - 1));
        const nm = String(dt.getMonth() + 1).padStart(2, '0');
        const nd = String(dt.getDate()).padStart(2, '0');
        const ny = d0.twoDigit ? String(dt.getFullYear()).slice(2) : String(dt.getFullYear());
        return `${nm}-${nd}-${ny}`;
      }
    }

    // Text with trailing number: "Item 1" → "Item 2"
    const trailRx = /^(.*?)(\\d+)$/;
    if (isSingle) {
      const tm = v.match(trailRx);
      if (tm) return `${tm[1]}${parseInt(tm[2], 10) + offsetSteps - 1}`;
    } else {
      const tm0 = srcVals[0].match(trailRx);
      const tm1 = srcVals[1]?.match(trailRx);
      if (tm0 && tm1 && tm0[1] === tm1[1]) {
        const step = parseInt(tm1[2], 10) - parseInt(tm0[2], 10);
        return `${tm0[1]}${parseInt(tm0[2], 10) + step * (offsetSteps - 1)}`;
      }
    }

    return v; // fallback: copy
  }"""

content = content.replace(old_apply, new_apply)

open(filepath, 'w', encoding='utf-8').write(content)
lines = content.splitlines()
print(f"Done. Total lines: {len(lines)}")
print("Checking key insertions...")
for kw in ['fillPopupState', 'executeFillMode', '_doFill', '_getNextSeriesEx', '_shiftFormula', 'fill-options-wrap', 'fill-options-dropdown']:
    count = content.count(kw)
    print(f"  {kw}: {count} occurrences")
