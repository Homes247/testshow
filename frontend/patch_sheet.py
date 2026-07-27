import re

with open(r"c:\Users\Homes247\Desktop\office-suite\frontend\src\app\pages\sheet-editor\sheet-editor.component.ts", "r", encoding="utf-8") as f:
    content = f.read()

# 1. Interface and state
if "export interface DropdownOption" not in content:
    content = content.replace("interface CellValidation {\n  type: 'list' | 'number' | 'date';\n  options?: string[];", "export interface DropdownOption {\n  label: string;\n  color?: string;\n}\n\nexport interface CellValidation {\n  type: 'list' | 'number' | 'date';\n  options?: (string | DropdownOption)[];")
    content = content.replace("validationInput = '';", "validationInput = '';\n  picklistOptions: DropdownOption[] = [];")

# 2. Add Helper Methods
helpers = """  getRangeRef(): string {
    if (this.rangeStart && this.rangeEnd) {
      const minR = Math.min(this.rangeStart.r, this.rangeEnd.r);
      const maxR = Math.max(this.rangeStart.r, this.rangeEnd.r);
      const minC = Math.min(this.rangeStart.c, this.rangeEnd.c);
      const maxC = Math.max(this.rangeStart.c, this.rangeEnd.c);
      if (minR === maxR && minC === maxC) return this.getCellRef(minR, minC);
      return `${this.getCellRef(minR, minC)}:${this.getCellRef(maxR, maxC)}`;
    }
    return this.getCellRef(this.selectedRow, this.selectedCol);
  }

  hasDropdownInRange(): boolean {
    if (!this.rangeStart || !this.rangeEnd) return this.hasCellDropdown(this.selectedRow, this.selectedCol);
    const minR = Math.min(this.rangeStart.r, this.rangeEnd.r);
    const maxR = Math.max(this.rangeStart.r, this.rangeEnd.r);
    const minC = Math.min(this.rangeStart.c, this.rangeEnd.c);
    const maxC = Math.max(this.rangeStart.c, this.rangeEnd.c);
    for (const key of Object.keys(this.validations)) {
      const parts = key.split(',');
      const r = parseInt(parts[0], 10);
      const c = parseInt(parts[1], 10);
      if (r >= minR && r <= maxR && c >= minC && c <= maxC) return true;
    }
    return false;
  }

  getCellDropdownOptions(r: number, c: number): (string | DropdownOption)[] {
    const v = this.validations[`${r},${c}`];
    return v && v.type === 'list' && v.options ? v.options : [];
  }

  getDropdownColor(r: number, c: number, val: string): string {
    const opts = this.getCellDropdownOptions(r, c);
    const found = opts.find(o => (typeof o === 'string' ? o : o.label) === val) as DropdownOption | undefined;
    return found?.color || '';
  }
"""
if "getRangeRef():" not in content:
    content = content.replace("  hasCellDropdown(r: number, c: number): boolean {\n    return !!this.validations[`${r},${c}`];\n  }", "  hasCellDropdown(r: number, c: number): boolean {\n    return !!this.validations[`${r},${c}`];\n  }\n\n" + helpers)

# 3. TextCell Template
old_textCell = r"""                <ng-template #textCell>
                  <ng-container *ngIf="hasCellDropdown(selectedRow, selectedCol) && r === selectedRow && c === selectedCol; else plainInput">
                    <select class="cell-input" [(ngModel)]="cells[r][c]" (change)="onCellChange(); save()" (click)="$event.stopPropagation()">
                      <option value=""></option>
                      <option *ngFor="let opt of validations[selectedRow + ',' + selectedCol].options" [value]="opt">{{opt}}</option>
                    </select>
                  </ng-container>"""
new_textCell = """                <ng-template #textCell>
                  <ng-container *ngIf="hasCellDropdown(r, c); else plainInput">
                    <select class="cell-select"
                      [(ngModel)]="cells[r][c]"
                      (focus)="selectCell(r, c)"
                      (change)="onCellChange(); save()"
                      (click)="$event.stopPropagation()"
                      [style.backgroundColor]="getDropdownColor(r, c, cells[r][c]) || 'transparent'"
                      [style.color]="getDropdownColor(r, c, cells[r][c]) ? '#fff' : 'inherit'"
                      style="appearance:none; border:none; outline:none; font-family:inherit; font-size:inherit; border-radius:0; padding:0 20px 0 4px; width:100%; height:100%; cursor:pointer; box-sizing:border-box; background-image: url('data:image/svg+xml;utf8,<svg fill=%22%23666%22 height=%2224%22 viewBox=%220 0 24 24%22 width=%2224%22 xmlns=%22http://www.w3.org/2000/svg%22><path d=%22M7 10l5 5 5-5z%22/></svg>'); background-repeat: no-repeat; background-position-x: calc(100% - 2px); background-position-y: center;">
                      <option value=""></option>
                      <option *ngFor="let opt of getCellDropdownOptions(r, c)" [value]="opt.label || opt" [style.background]="opt.color || '#fff'" [style.color]="opt.color ? '#fff' : '#000'">
                        {{ opt.label || opt }}
                      </option>
                    </select>
                  </ng-container>"""
content = content.replace(old_textCell, new_textCell)

# 4. Context Menu
old_ctx = r"""        <div class="ctx-item danger" *ngIf="hasCellDropdown(selectedRow, selectedCol)" (click)="removeValidation(); hideCtx()"><span class="ctx-icon">🗑️</span> Remove dropdown</div>"""
new_ctx = """        <div class="ctx-item danger" *ngIf="hasDropdownInRange()" (click)="removeValidation(); hideCtx()"><span class="ctx-icon" style="font-size: 16px;">&times;</span> Remove dropdown</div>"""
content = content.replace(old_ctx, new_ctx)

old_th = r"""<th *ngFor="let c of colRange" class="col-head"
                [class.col-selected]="isColHeaderSelected(c)"
                (click)="selectEntireCol(c)">{{ colLabel(c) }}</th>"""
new_th = """<th *ngFor="let c of colRange" class="col-head"
                [class.col-selected]="isColHeaderSelected(c)"
                (contextmenu)="onHeaderRightClick($event, 'col', c)"
                (click)="selectEntireCol(c)">{{ colLabel(c) }}</th>"""
content = content.replace(old_th, new_th)

old_td = r"""<td class="row-head" [class.row-selected]="isRowHeaderSelected(r)" (click)="selectEntireRow(r)">{{ r + 1 }}</td>"""
new_td = """<td class="row-head" [class.row-selected]="isRowHeaderSelected(r)" (contextmenu)="onHeaderRightClick($event, 'row', r)" (click)="selectEntireRow(r)">{{ r + 1 }}</td>"""
content = content.replace(old_td, new_td)

if "onHeaderRightClick(" not in content:
    header_ctx_code = """  onHeaderRightClick(e: MouseEvent, type: 'row'|'col', idx: number) {
    e.preventDefault();
    if (type === 'col') this.selectEntireCol(idx);
    else this.selectEntireRow(idx);
    const menuWidth = 220;
    const menuHeight = 500;
    let x = e.clientX;
    let y = e.clientY;
    if (x + menuWidth > window.innerWidth) x = window.innerWidth - menuWidth;
    if (y + menuHeight > window.innerHeight) y = Math.max(0, window.innerHeight - menuHeight);
    this.ctxX = x;
    this.ctxY = y;
    this.ctxVisible = true;
  }
"""
    content = content.replace("  onCellRightClick(e: MouseEvent, r: number, c: number) {", header_ctx_code + "\n  onCellRightClick(e: MouseEvent, r: number, c: number) {")

# 5. Validation Modal
old_modal = r"""      <!-- Validation Modal -->
      <div class="modal-overlay" *ngIf="validationModalOpen" (click)="validationModalOpen = false">
        <div class="modal" (click)="$event.stopPropagation()">
          <h3 style="margin-top:0;">Data Validation</h3>
          <p style="color:#5f6368;font-size:13px;margin-bottom:12px;">Applies to cell {{getCellRef(selectedRow, selectedCol)}}</p>
          <label style="display:block;font-size:12px;font-weight:600;margin-bottom:6px;color:#5f6368;">Dropdown List Items</label>
          <textarea [(ngModel)]="validationInput" style="width:100%;height:80px;padding:8px;border:1px solid #dadce0;border-radius:4px;resize:vertical;" 
            placeholder="e.g.&#10;Closed&#10;New&#10;Fixed&#10;Reopened"></textarea>
          <div style="display:flex;gap:10px;margin-top:14px;justify-content:flex-end;">
            <button class="btn outline" (click)="validationModalOpen = false">Cancel</button>
            <button class="btn" (click)="saveValidation()">Apply</button>
          </div>
        </div>
      </div>"""
new_modal = """      <!-- Validation / Dropdown Modal (Zoho Picklist Style) -->
      <div class="modal-overlay" *ngIf="validationModalOpen" (click)="validationModalOpen = false">
        <div class="modal" (click)="$event.stopPropagation()" style="width:360px; background:#1c2333; color:#fff; border:1px solid #2d3748; box-shadow:0 12px 40px rgba(0,0,0,0.5); padding:20px;">
          <button (click)="validationModalOpen = false" style="position:absolute;top:16px;right:16px;background:none;border:none;font-size:18px;cursor:pointer;color:#a0aec0;display:flex;align-items:center;justify-content:center;width:24px;height:24px;border-radius:50%;background:rgba(255,255,255,0.1);"><span class="material-symbols-outlined" style="font-size:16px;">close</span></button>
          <h3 style="margin-top:0;font-size:16px;font-weight:600;margin-bottom:16px;">Picklist - Edit</h3>
          <p style="color:#a0aec0;font-size:12px;margin-bottom:16px;display:flex;align-items:center;gap:4px;">
            Applies to: <span style="color:#81e6d9;">{{getRangeRef()}}</span>
          </p>
          <div style="max-height:240px;overflow-y:auto;margin-bottom:16px;padding-right:4px;">
            <div *ngFor="let opt of picklistOptions; let i = index" style="display:flex; align-items:center; gap:8px; margin-bottom:8px;">
               <div style="position:relative; width:24px; height:24px; border-radius:4px; overflow:hidden; border:1px solid rgba(255,255,255,0.2); flex-shrink:0;">
                 <input type="color" [(ngModel)]="opt.color" style="position:absolute;top:-5px;left:-5px;width:40px;height:40px;border:none;cursor:pointer;padding:0;background:transparent;">
               </div>
               <input type="text" [(ngModel)]="opt.label" placeholder="Option label" style="flex:1; background:#2d3748; border:1px solid transparent; color:#fff; padding:6px 10px; border-radius:4px; outline:none; font-size:13px; transition:border 0.2s;">
               <button (click)="picklistOptions.splice(i, 1)" style="background:none;border:none;color:#fc8181;cursor:pointer;font-size:18px;padding:4px;display:flex;align-items:center;"><span class="material-symbols-outlined" style="font-size:18px;">close</span></button>
            </div>
          </div>
          <button (click)="addPicklistOption()" style="background:none;border:none;color:#81e6d9;cursor:pointer;font-weight:500;font-size:13px;padding:0;display:flex;align-items:center;gap:4px;margin-bottom:24px;">
            <span class="material-symbols-outlined" style="font-size:16px;">add_circle</span> Add New
          </button>
          <div style="display:flex;gap:10px;justify-content:flex-end;">
            <button (click)="validationModalOpen = false" style="background:transparent;border:1px solid rgba(255,255,255,0.2);color:#fff;padding:8px 16px;border-radius:6px;cursor:pointer;font-size:13px;">Cancel</button>
            <button (click)="saveValidation()" style="background:#00c274;border:none;color:#fff;padding:8px 16px;border-radius:6px;cursor:pointer;font-weight:600;font-size:13px;">Save</button>
          </div>
        </div>
      </div>"""
content = content.replace(old_modal, new_modal)

# 6. saveValidation and removeValidation methods
old_methods = r"""  openValidationModal() {
    this.validationInput = '';
    const existing = this.validations[`${this.selectedRow},${this.selectedCol}`];
    if (existing && existing.options) {
      this.validationInput = existing.options.join('\n');
    }
    this.validationModalOpen = true;
  }

  saveValidation() {
    const raw = this.validationInput.trim();
    if (!raw) { this.validationModalOpen = false; return; }
    // Accept newline or comma separated options
    const options = raw.split(/[\n,]/).map(o => o.trim()).filter(o => o.length > 0);
    if (options.length === 0) { this.validationModalOpen = false; return; }
    this.validations = {
      ...this.validations,
      [`${this.selectedRow},${this.selectedCol}`]: { type: 'list', options }
    };
    // If cell has a value not in options, clear it
    const cur = this.cells[this.selectedRow][this.selectedCol];
    if (cur && !options.includes(cur)) this.cells[this.selectedRow][this.selectedCol] = '';
    this.validationModalOpen = false;
    this.onCellChange();
    this.save();
    this.showToast(`Dropdown set: ${options.length} options`);
  }

  removeValidation() {
    const key = `${this.selectedRow},${this.selectedCol}`;
    if (this.validations[key]) {
      const v = { ...this.validations };
      delete v[key];
      this.validations = v;
      this.onCellChange();
      this.save();
      this.showToast('Dropdown removed.');
    }
  }"""
new_methods = """  openValidationModal() {
    const existing = this.validations[`${this.selectedRow},${this.selectedCol}`];
    this.picklistOptions = [];
    if (existing && existing.options) {
      existing.options.forEach(o => {
        if (typeof o === 'string') this.picklistOptions.push({ label: o, color: '#4a5568' });
        else this.picklistOptions.push({ label: (o as DropdownOption).label, color: (o as DropdownOption).color || '#4a5568' });
      });
    } else {
      this.picklistOptions.push({ label: 'Item 1', color: '#4caf50' });
      this.picklistOptions.push({ label: 'Item 2', color: '#f44336' });
    }
    this.validationModalOpen = true;
  }

  addPicklistOption() {
    const colors = ['#4caf50', '#f44336', '#ff9800', '#2196f3', '#9c27b0', '#795548', '#607d8b'];
    this.picklistOptions.push({ label: '', color: colors[this.picklistOptions.length % colors.length] });
  }

  saveValidation() {
    const validOptions = this.picklistOptions.filter(o => o.label.trim().length > 0);
    if (validOptions.length === 0) { this.validationModalOpen = false; return; }
    let minR = this.selectedRow, maxR = this.selectedRow;
    let minC = this.selectedCol, maxC = this.selectedCol;
    if (this.rangeStart && this.rangeEnd) {
      minR = Math.min(this.rangeStart.r, this.rangeEnd.r);
      maxR = Math.max(this.rangeStart.r, this.rangeEnd.r);
      minC = Math.min(this.rangeStart.c, this.rangeEnd.c);
      maxC = Math.max(this.rangeStart.c, this.rangeEnd.c);
    }
    const newValidations = { ...this.validations };
    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        newValidations[`${r},${c}`] = { type: 'list', options: validOptions };
        const cur = this.cells[r][c];
        if (cur && !validOptions.find(o => o.label === cur)) this.cells[r][c] = '';
      }
    }
    this.validations = newValidations;
    this.validationModalOpen = false;
    this.onCellChange();
    this.save();
    this.showToast(`Picklist set: ${validOptions.length} items`);
  }

  removeValidation() {
    let minR = this.selectedRow, maxR = this.selectedRow;
    let minC = this.selectedCol, maxC = this.selectedCol;
    if (this.rangeStart && this.rangeEnd) {
      minR = Math.min(this.rangeStart.r, this.rangeEnd.r);
      maxR = Math.max(this.rangeStart.r, this.rangeEnd.r);
      minC = Math.min(this.rangeStart.c, this.rangeEnd.c);
      maxC = Math.max(this.rangeStart.c, this.rangeEnd.c);
    }
    let removed = false;
    const v = { ...this.validations };
    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        const key = `${r},${c}`;
        if (v[key]) {
          delete v[key];
          removed = true;
        }
      }
    }
    if (removed) {
      this.validations = v;
      this.onCellChange();
      this.save();
      this.showToast('Dropdown removed.');
    }
  }"""
content = content.replace(old_methods, new_methods)

# 7. applyFill update
old_applyFill = r"""  private applyFill() {
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
            const idx = (r - srcMaxR - 1) % srcVals.length;
            this.cells[r][c] = this.getNextSeriesValue(srcVals, r - srcMinR);
          }
        } else {
          for (let r = srcMinR - 1; r >= dstR; r--) {
            this.cells[r][c] = this.getNextSeriesValue(srcVals, r - srcMaxR);
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
            this.cells[r][c] = this.getNextSeriesValue(srcVals, c - srcMinC);
          }
        } else {
          for (let c = srcMinC - 1; c >= dstC; c--) {
            this.cells[r][c] = this.getNextSeriesValue(srcVals, c - srcMaxC);
          }
        }
      }
    }

    // Update range
    if (goDown) this.rangeEnd = { r: dstR, c: srcMaxC };
    if (goUp) this.rangeStart = { r: dstR, c: srcMinC };
    if (goRight) this.rangeEnd = { r: srcMaxR, c: dstC };
    if (goLeft) this.rangeStart = { r: srcMinR, c: dstC };

    this.onCellChange();
    this.save();
  }"""
new_applyFill = """  private applyFill() {
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

    const newValidations = { ...this.validations };

    if (goDown || goUp) {
      // Fill column-wise
      for (let c = srcMinC; c <= srcMaxC; c++) {
        const srcVals: string[] = [];
        for (let r = srcMinR; r <= srcMaxR; r++) srcVals.push(this.cells[r][c]);
        const rangeLen = srcMaxR - srcMinR + 1;
        if (goDown) {
          for (let r = srcMaxR + 1; r <= dstR; r++) {
            const idx = (r - srcMaxR - 1) % rangeLen;
            const srcR = srcMinR + idx;
            this.cells[r][c] = this.getNextSeriesValue(srcVals, r - srcMinR);
            if (this.validations[`${srcR},${c}`]) newValidations[`${r},${c}`] = { ...this.validations[`${srcR},${c}`] };
            else delete newValidations[`${r},${c}`];
          }
        } else {
          for (let r = srcMinR - 1; r >= dstR; r--) {
            const idx = (srcMinR - r - 1) % rangeLen;
            const srcR = srcMaxR - idx;
            this.cells[r][c] = this.getNextSeriesValue(srcVals, r - srcMaxR);
            if (this.validations[`${srcR},${c}`]) newValidations[`${r},${c}`] = { ...this.validations[`${srcR},${c}`] };
            else delete newValidations[`${r},${c}`];
          }
        }
      }
    } else if (goRight || goLeft) {
      // Fill row-wise
      for (let r = srcMinR; r <= srcMaxR; r++) {
        const srcVals: string[] = [];
        for (let c = srcMinC; c <= srcMaxC; c++) srcVals.push(this.cells[r][c]);
        const rangeLen = srcMaxC - srcMinC + 1;
        if (goRight) {
          for (let c = srcMaxC + 1; c <= dstC; c++) {
            const idx = (c - srcMaxC - 1) % rangeLen;
            const srcC = srcMinC + idx;
            this.cells[r][c] = this.getNextSeriesValue(srcVals, c - srcMinC);
            if (this.validations[`${r},${srcC}`]) newValidations[`${r},${c}`] = { ...this.validations[`${r},${srcC}`] };
            else delete newValidations[`${r},${c}`];
          }
        } else {
          for (let c = srcMinC - 1; c >= dstC; c--) {
            const idx = (srcMinC - c - 1) % rangeLen;
            const srcC = srcMaxC - idx;
            this.cells[r][c] = this.getNextSeriesValue(srcVals, c - srcMaxC);
            if (this.validations[`${r},${srcC}`]) newValidations[`${r},${c}`] = { ...this.validations[`${r},${srcC}`] };
            else delete newValidations[`${r},${c}`];
          }
        }
      }
    }
    
    this.validations = newValidations;

    // Update range
    if (goDown) this.rangeEnd = { r: dstR, c: srcMaxC };
    if (goUp) this.rangeStart = { r: dstR, c: srcMinC };
    if (goRight) this.rangeEnd = { r: srcMaxR, c: dstC };
    if (goLeft) this.rangeStart = { r: srcMinR, c: dstC };

    this.onCellChange();
    this.save();
  }"""
content = content.replace(old_applyFill, new_applyFill)

with open(r"c:\Users\Homes247\Desktop\office-suite\frontend\src\app\pages\sheet-editor\sheet-editor.component.ts", "w", encoding="utf-8") as f:
    f.write(content)

print("Patch applied successfully.")
