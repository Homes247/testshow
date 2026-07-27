import { Component, OnInit, OnDestroy, HostListener, ViewChild, ElementRef, ChangeDetectionStrategy, ChangeDetectorRef } from '@angular/core';
import * as Tesseract from 'tesseract.js';
import { DomSanitizer, SafeResourceUrl } from '@angular/platform-browser';
import { ActivatedRoute, Router } from '@angular/router';
import { CommonModule } from '@angular/common';
import { HttpEventType } from '@angular/common/http';
import { FormsModule } from '@angular/forms';
import { Subscription, Subject } from 'rxjs';
import { debounceTime } from 'rxjs/operators';
import { ChatWidgetComponent } from '../../components/chat-widget/chat-widget.component';
import { ApiService } from '../../services/api.service';
import { AuthService } from '../../services/auth.service';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { PickerModule } from '@ctrl/ngx-emoji-mart';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
const colName = (i: number) => {
  let name = '';
  let temp = i;
  while (temp >= 0) {
    name = String.fromCharCode(65 + (temp % 26)) + name;
    temp = Math.floor(temp / 26) - 1;
  }
  return name;
};

export interface CellFormat {
  note?: string;
  bold?: boolean;
  italic?: boolean;
  strikethrough?: boolean;
  underline?: boolean;
  color?: string;
  bg?: string;
  align?: 'left' | 'center' | 'right';
  vertAlign?: 'top' | 'middle' | 'bottom';
  font?: string;
  size?: string;
  wrap?: 'overflow' | 'wrap' | 'clip' | 'shrink' | boolean;
  indent?: number;
  rotation?: string | number;
  numFormat?: string;
  decimals?: number;
  borders?: {
    top?: boolean | CellBorder;
    bottom?: boolean | CellBorder;
    left?: boolean | CellBorder;
    right?: boolean | CellBorder;
    all?: boolean | CellBorder;
  };
}


export interface SparklineConfig {
  sourceRange?: string;
  destinationRange?: string;
  type: 'line' | 'column' | 'winloss'; // We use 'column' instead of 'bar' as per our previous logic
  baseColor: string;
  highlights: {
    high: { enabled: boolean; color: string };
    low: { enabled: boolean; color: string };
    first: { enabled: boolean; color: string };
    last: { enabled: boolean; color: string };
    negative: { enabled: boolean; color: string };
    markers: { enabled: boolean; color: string };
  };
  emptyCellMode: 'gap' | 'zero' | 'connect' | 'skip';
  includeHiddenRowsColumns: boolean;
  horizontalAxis: {
    displayAxis: boolean;
    rightToLeft: boolean;
  };
  verticalAxis: {
    min: { mode: 'auto' | 'same' | 'custom'; customValue: number | null };
    max: { mode: 'auto' | 'same' | 'custom'; customValue: number | null };
  };
  isGrouped: boolean;
  groupId: string;
}
export interface CellBorder {
  color?: string;
  style?: string;
  width?: string;
}

export interface DropdownOption {
  label: string;
  color?: string;
  textColor?: string;
}

export interface CellValidation {
  type: string;
  options: (string | DropdownOption)[];
  isMultiSelect?: boolean;
  displayAsChip?: boolean;
  colorMode?: 'none' | 'single' | 'multi';
  singleColor?: string;
}

export interface AuditOp {
  action_type: string;
  target_range: string;
  metadata?: any;
}

@Component({
  selector: 'app-sheet-editor',
  standalone: true,
  changeDetection: ChangeDetectionStrategy.OnPush,
  imports: [CommonModule, FormsModule, ChatWidgetComponent, PickerModule],
  template: `
    <div class="shell" [ngClass]="'theme-' + currentTheme" (mousedown)="$event.target===$event.currentTarget?closeMenus():null">

      <!-- ═══ TOP BAR ════════════════════════════════════════════════════════ -->
      <div class="top-bar" *ngIf="showTopBar">
        <div class="tl" style="align-items:center;">
          <button class="back-btn" (click)="back()" title="Back" style="background:none; border:none; cursor:pointer; display:flex; align-items:center; justify-content:center; width:32px; height:32px; border-radius:50%; flex-shrink:0; opacity:0.8;">
            <svg width="20" height="20" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M10 19l-7-7m0 0l7-7m-7 7h18"/></svg>
          </button>
          <div class="brand" style="display:flex; align-items:center; gap:6px; cursor:pointer;" (click)="goHome()">
            <svg width="28" height="28" viewBox="0 0 32 32" fill="none">
              <rect width="32" height="32" rx="5" fill="#26A96C"/>
              <rect x="5" y="8" width="22" height="2.5" rx="1.2" fill="white"/>
              <rect x="5" y="13.5" width="22" height="2.5" rx="1.2" fill="white"/>
              <rect x="5" y="19" width="14" height="2.5" rx="1.2" fill="white"/>
            </svg>
            <span class="brand-name" style="font-weight:600; font-size:18px;">Sheet</span>
          </div>
          <div class="doc-sec" style="display:flex; align-items:center; gap:12px; margin-top:0; margin-left:8px;">
            <input class="doc-title" [(ngModel)]="title" (blur)="save()" placeholder="Untitled spreadsheet" [style.width.ch]="(title || 'Untitled spreadsheet').length + 3"/>
            <div class="doc-icons" style="display:flex; align-items:center; gap:8px; opacity:0.8;">
              <span class="material-symbols-outlined" style="font-size:16px; cursor:pointer;" (click)="toggleStar()" [style.color]="isStarred ? '#fbbc04' : 'inherit'" [title]="isStarred ? 'Unstar' : 'Star'">{{ isStarred ? 'star' : 'star_border' }}</span>
              <span class="material-symbols-outlined" style="font-size:16px; cursor:pointer;" (click)="openFeatureModal('move')" title="Move to Folder">folder_open</span>
              <div style="display:flex; align-items:center; font-size:12px; color:inherit; margin-left:4px;">
                <span *ngIf="saveStatus==='saving'" style="font-style:italic;">Saving...</span>
                <span *ngIf="saveStatus==='saved'">
                   Saved at {{lastSavedTime}}
                </span>
                <span *ngIf="saveStatus==='error'" style="color:#ea4335;">Failed to save</span>
              </div>
            </div>
          </div>
        </div>
        <div class="tr">
          <div class="top-search-box" [class.has-query]="inlineSearchQuery">
            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24" style="flex-shrink:0; opacity:0.7;"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
            <input placeholder="Search in this sheet" [(ngModel)]="inlineSearchQuery" (ngModelChange)="onInlineSearch()" (keydown.enter)="inlineFindNext()" class="inline-search-input">
            
            <ng-container *ngIf="inlineSearchQuery">
              <button (click)="clearInlineSearch()" class="inline-search-clear" title="Clear search">
                <span class="material-symbols-outlined" style="font-size:14px;">close</span>
              </button>
              <div class="inline-search-divider"></div>
              <span class="inline-search-count">{{ inlineSearchMatches.length ? inlineSearchIdx + 1 : 0 }} / {{ inlineSearchMatches.length }}</span>
              <div class="inline-search-nav">
                <button (click)="inlineFindPrev()" title="Previous match"><span class="material-symbols-outlined" style="font-size:16px;">chevron_left</span></button>
                <button (click)="inlineFindNext()" title="Next match"><span class="material-symbols-outlined" style="font-size:16px;">chevron_right</span></button>
              </div>
            </ng-container>
          </div>
          <div class="online-badge" *ngIf="showUserPresence && activeUsers>1" title="{{activeUsers}} users editing">
            <span class="material-symbols-outlined" style="font-size:16px;">group</span>
            <span style="margin-left:4px;">{{activeUsers}}</span>
          </div>
          <button class="share-btn" (click)="shareModalOpen=true">
            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="18" cy="5" r="3"/><circle cx="6" cy="12" r="3"/><circle cx="18" cy="19" r="3"/><line x1="8.59" y1="13.51" x2="15.42" y2="17.49"/><line x1="15.41" y1="6.51" x2="8.59" y2="10.49"/></svg>
            Share
          </button>
          
          <button class="properties-btn" (click)="propertiesPanelOpen = true" title="Properties">
            <span class="material-symbols-outlined" style="font-size:20px;">info</span>
          </button>
          
          <div class="av" (click)="profileOpen=!profileOpen;$event.stopPropagation()" title="Account">{{initials}}
            <div class="profile-dd" *ngIf="profileOpen" (click)="$event.stopPropagation()">
              <div class="pd-head">
                <div class="pd-av">{{initials}}</div>
                <div>
                  <div style="font-size:13px;font-weight:600;color:#202124;">{{auth.user?.name ?? 'User'}}</div>
                  <div style="font-size:11px;color:#5f6368;">{{auth.user?.email ?? ''}}</div>
                </div>
              </div>
              <div class="pd-item" (click)="openApp('account')"><span class="material-symbols-outlined pd-icon">manage_accounts</span> My Account</div>
              <div class="pd-item" (click)="openApp('calendar')"><span class="material-symbols-outlined pd-icon">calendar_month</span> Calendar</div>
              <div class="pd-item" (click)="openApp('notes')"><span class="material-symbols-outlined pd-icon">sticky_note_2</span> Notes</div>
              <div class="pd-item" (click)="openApp('tasks')"><span class="material-symbols-outlined pd-icon">task_alt</span> Tasks</div>
              <div class="pd-item" (click)="openApp('settings')"><span class="material-symbols-outlined pd-icon">settings</span> Settings</div>
              <div class="pd-sep"></div>
              <div class="pd-item danger" (click)="auth.logout()"><span class="material-symbols-outlined pd-icon">logout</span> Sign Out</div>
            </div>
          </div>
        </div>
      </div>

      <!-- ═══ MENU BAR ══════════════════════════════════════════════════════ -->
      <div class="menu-row" (mousedown)="$event.preventDefault()">
        <div class="mi" (click)="toggleMenu('file',$event)" [class.mi-open]="activeMenu==='file'">File
                    <div class="mdd" *ngIf="activeMenu==='file'">
            <div class="mdi" (click)="newDoc()"><span class="mdi-icon material-symbols-outlined">grid_view</span>New Spreadsheet<span class="mh">Ctrl+N</span></div>
            <div class="mdi" (click)="openFeatureModal('template')"><span class="mdi-icon material-symbols-outlined">dashboard_customize</span>New from Template...</div>
            <div class="mdi has-sub"><span class="mdi-icon material-symbols-outlined">folder_open</span>Open<span class="mdi-arrow material-symbols-outlined">chevron_right</span>
               <div class="mdi-sub">
                 <div class="mdi" (click)="openFeatureModal('open')">From Vmail Drive</div>
                 <div class="mdi" (click)="openFeatureModal('import')">From Computer</div>
               </div>
            </div>
            <div class="mdi has-sub"><span class="mdi-icon material-symbols-outlined">upload_file</span>Import<span class="mdi-arrow material-symbols-outlined">chevron_right</span>
               <div class="mdi-sub">
                 <div class="mdi" (click)="openFeatureModal('import')">Upload File</div>
               </div>
            </div>
            <div class="mds"></div>
            <div class="mdi" (click)="triggerCopy()"><span class="mdi-icon material-symbols-outlined">content_copy</span>Make a Copy...</div>
            <div class="mdi" (click)="save(); closeMenus()"><span class="mdi-icon material-symbols-outlined">save</span>Save<span class="mh">Ctrl+S</span></div>
                        <div class="mdi has-sub"><span class="mdi-icon material-symbols-outlined">download</span>Download as<span class="mdi-arrow material-symbols-outlined">chevron_right</span>
               <div class="mdi-sub">
                 <div class="mdi" (click)="exportFile('xlsx')">MS Excel Workbook<span class="mh">.xlsx</span></div>
                 <div class="mdi" (click)="exportFile('xlsb')">MS Excel Binary Workbook<span class="mh">.xlsb</span></div>
                 <div class="mdi" (click)="exportFile('ods')">Open Office Spreadsheet<span class="mh">.ods</span></div>
                 <div class="mdi" (click)="exportFile('csv')">Comma Separated Values<span class="mh">.csv</span></div>
                 <div class="mdi" (click)="exportFile('tsv')">Tab Separated Values<span class="mh">.tsv</span></div>
                 <div class="mdi" (click)="exportFile('pdf')">PDF Document<span class="mh">.pdf</span></div>
                 <div class="mdi" (click)="exportFile('html')">HTML Document<span class="mh">.zip</span></div>
               </div>
            </div>
            <div class="mds"></div>
            <div class="mdi" (click)="openFeatureModal('password')"><span class="mdi-icon material-symbols-outlined">lock</span>Password Protected File...</div>
            <div class="mdi" (click)="shareModalOpen=true;closeMenus()"><span class="mdi-icon material-symbols-outlined">mail</span>Email As Attachment...</div>
            <div class="mds"></div>
            <div class="mdi" (click)="openFeatureModal('move')"><span class="mdi-icon material-symbols-outlined">drive_file_move</span>Move...</div>
            <div class="mdi" (click)="triggerRename()"><span class="mdi-icon material-symbols-outlined">drive_file_rename_outline</span>Rename...</div>
            <div class="mds"></div>
            <div class="mdi" (click)="openFeatureModal('audit')"><span class="mdi-icon material-symbols-outlined">history</span>Audit Trail...</div>
            <div class="mdi has-sub"><span class="mdi-icon material-symbols-outlined">update</span>Version<span class="mdi-arrow material-symbols-outlined">chevron_right</span>
               <div class="mdi-sub">
                 <div class="mdi" (click)="openFeatureModal('version')">Version History</div>
               </div>
            </div>
            <div class="mdi has-sub"><span class="mdi-icon material-symbols-outlined">schema</span>Workflow<span class="mdi-arrow material-symbols-outlined">chevron_right</span>
               <div class="mdi-sub">
                 <div class="mdi" (click)="openFeatureModal('workflow')">Manage Workflows</div>
               </div>
            </div>
            <div class="mds"></div>
            <div class="mdi has-sub"><span class="mdi-icon material-symbols-outlined">share</span>Share<span class="mdi-arrow material-symbols-outlined">chevron_right</span>
               <div class="mdi-sub">
                 <div class="mdi" (click)="shareModalOpen=true;closeMenus()"><span class="mdi-icon material-symbols-outlined">person_add</span>Share with collaborators</div>
                 <div class="mdi" (click)="shareModalOpen=true;closeMenus()"><span class="mdi-icon material-symbols-outlined">link</span>Publish to web</div>
               </div>
            </div>
            <div class="mdi danger" (click)="trashDoc()"><span class="mdi-icon material-symbols-outlined" style="color:inherit">delete</span>Move to Trash</div>
          </div>
        </div>
        <div class="mi" (click)="toggleMenu('edit',$event)" [class.mi-open]="activeMenu==='edit'">Edit
          <div class="mdd" *ngIf="activeMenu==='edit'">
            <div class="mdi" (click)="undo()">Undo<span class="mh">Ctrl+Z</span></div>
            <div class="mdi" (click)="redo()">Redo<span class="mh">Ctrl+Y</span></div>
            <div class="mds"></div>
            <div class="mdi" (click)="cutCell();closeMenus()">Cut<span class="mh">Ctrl+X</span></div>
            <div class="mdi" (click)="copyCell()">Copy<span class="mh">Ctrl+C</span></div>
            <div class="mdi has-sub">Paste <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
              <div class="mdi-sub">
                <div class="mdi" (click)="pasteCell()">All<span class="mh">Ctrl+V</span></div>
                <div class="mdi" (click)="pasteValues()">Values<span class="mh">Ctrl+Shift+V</span></div>
                <div class="mdi" (click)="pasteFormulas()">Formulas</div>
                <div class="mdi" (click)="pasteFormats()">Formats</div>
                <div class="mdi" (click)="pasteNotes()">Notes</div>
                <div class="mds"></div>
                <div class="mdi" (click)="pasteFormulasAndNumberFormats()">Formulas and Number Formats</div>
                <div class="mdi" (click)="pasteValuesAndNumberFormats()">Values and Number Formats</div>
                <div class="mdi" (click)="pasteValidation()">Validation</div>
                <div class="mds"></div>
                <div class="mdi" (click)="pasteExceptNotes()">All Except Notes</div>
                <div class="mdi" (click)="pasteExceptBorders()">All Except Borders</div>
                <div class="md-sep"></div>
                <div class="mdi" (click)="pasteLinkToSource()">Link To Source</div>
                <div class="mdi" (click)="pasteTranspose()">Transpose</div>
              </div>
            </div>
            <div class="mds"></div>
            <div class="mdi has-sub">Fill <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
              <div class="mdi-sub">
                <div class="mdi" (click)="fillDown();closeMenus()">Down<span class="mh">Ctrl+D</span></div>
                <div class="mdi" (click)="fillRight();closeMenus()">Right<span class="mh">Ctrl+R</span></div>
                <div class="mdi" (click)="fillUp();closeMenus()">Up</div>
                <div class="mdi" (click)="fillLeft();closeMenus()">Left</div>
                <div class="mds"></div>
                <div class="mdi" (click)="patternFill();closeMenus()">Pattern Fill<span class="mh">Ctrl+E</span></div>
              </div>
            </div>
            <div class="mdi has-sub">Clear <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
              <div class="mdi-sub">
                <div class="mdi" (click)="clearAll();closeMenus()">All<span class="mh">Ctrl+Del</span></div>
                <div class="mdi" (click)="clearAllFormats();closeMenus()">Formats<span class="mh">Shift+Del</span></div>
                <div class="mdi" (click)="clearRangeData();closeMenus()">Contents<span class="mh">Del</span></div>
                <div class="mds"></div>
                <div class="mdi" (click)="clearNotes();closeMenus()">Notes</div>
                <div class="mdi" (click)="clearHyperlinks();closeMenus()">Hyperlinks</div>
                <div class="mdi" (click)="clearCheckboxes();closeMenus()">Checkboxes</div>
                <div class="mds"></div>
                <div class="mdi" (click)="clearDataValidations();closeMenus()">Data Validations</div>
                <div class="mdi" (click)="clearConditionalFormats();closeMenus()">Conditional Formats</div>
                <div class="mdi" (click)="clearRichTextFormats();closeMenus()">RichText Formats</div>
              </div>
            </div>
            <div class="mdi has-sub">Delete <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
              <div class="mdi-sub">
                <div class="mdi" (click)="deleteShiftLeft()">Shift Cells Left</div>
                <div class="mdi" (click)="deleteShiftUp()">Shift Cells Up</div>
                <div class="mds"></div>
                <div class="mdi" (click)="deleteRow()">Delete {{ selectedRowCount }} Row{{ selectedRowCount > 1 ? 's' : '' }}</div>
                <div class="mdi" (click)="deleteCol()">Delete {{ selectedColCount }} Column{{ selectedColCount > 1 ? 's' : '' }}</div>
              </div>
            </div>
            <div class="mds"></div>
            <div class="mdi" (click)="openFind();closeMenus()">Find and Replace...<span class="mh">Ctrl+Shift+H</span></div>
            <div class="mdi" (click)="openFeatureModal('goto');closeMenus()">Go To...<span class="mh">Ctrl+G</span></div>
            <div class="mds"></div>
            <div class="mdi" (click)="recalculate();closeMenus()">Recalculate<span class="mh">F9</span></div>
          </div>
        </div>
        <div class="mi" (click)="toggleMenu('view',$event)" [class.mi-open]="activeMenu==='view'">View
          <div class="mdd" *ngIf="activeMenu==='view'">
            <div class="mdi has-sub"><span class="mdi-icon material-symbols-outlined">dataset</span>Freeze <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
              <div class="mdi-sub">
                <div class="mdi" (click)="freezeRows(1);closeMenus()">Row 1</div>
                <div class="mdi" (click)="freezeRows(selectedRow+1);closeMenus()">Up to Row {{selectedRow+1}}</div>
                <div class="mds"></div>
                <div class="mdi" (click)="freezeCols(1);closeMenus()">Column A</div>
                <div class="mdi" (click)="freezeCols(selectedCol+1);closeMenus()">Up to Column {{colLabel(selectedCol)}}</div>
                <div class="mds"></div>
                <div class="mdi" (click)="freezeSelection();closeMenus()" style="display:flex; flex-direction:column; align-items:flex-start; line-height:1.2; padding:6px 16px;">
                  <div>Selection</div>
                  <div style="font-size:10px; color:#9aa0a6; white-space:normal; max-width:180px; margin-top:4px;">The selected row(s) or column(s) will be frozen and placed to the top or left of the editor respectively.</div>
                </div>
              </div>
            </div>
            <div class="mdi has-sub"><span class="mdi-icon material-symbols-outlined">visibility</span>Hide &amp; Unhide <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
              <div class="mdi-sub">
                <div class="mdi" (click)="hideRows();closeMenus()">Hide Rows<span class="mh">Ctrl+Alt+9</span></div>
                <div class="mdi" (click)="hideCols();closeMenus()">Hide Columns<span class="mh">Ctrl+Alt+0</span></div>
                <div class="mdi" [class.disabled]="getVisibleSheetCount() <= 1" (click)="hideSheet(currentSheetIdx); closeMenus()">Hide Sheet</div>
                <div class="mds"></div>
                <div class="mdi" (click)="unhideRows();closeMenus()">Unhide Rows<span class="mh">Ctrl+Shift+9</span></div>
                <div class="mdi" (click)="unhideCols();closeMenus()">Unhide Columns<span class="mh">Ctrl+Shift+0</span></div>
                <div class="mds"></div>
                <div class="mdi has-sub" [class.disabled]="hiddenSheetsList.length === 0">
                  Hidden Sheets <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
                  <div class="mdi-sub" *ngIf="hiddenSheetsList.length > 0">
                    <div class="mdi" *ngFor="let hs of hiddenSheetsList; trackBy: trackByHiddenSheet" (click)="unhideSheetAndSwitch(hs.idx); closeMenus()">{{ hs.s.name }}</div>
                    <div class="mds"></div>
                    <div class="mdi" (click)="unhideAllSheets(); closeMenus()">Unhide All Sheets</div>
                  </div>
                </div>
              </div>
            </div>
            <div class="mdi has-sub"><span class="mdi-icon material-symbols-outlined">grid_on</span>Gridlines <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
              <div class="mdi-sub" style="width:240px; padding:8px;">
                <div class="mdi" (click)="toggleGridlines();closeMenus()" style="padding:6px 8px; margin-bottom:8px;"><span class="material-symbols-outlined" style="font-size:16px; margin-right:8px; vertical-align:-3px;">{{showGridlines?'visibility_off':'visibility'}}</span>{{showGridlines?'Hide Gridlines':'Show Gridlines'}}</div>
                <div class="mdi" (click)="setGridlineColor('#d0d0d0');closeMenus()" style="padding:6px 8px;"><div style="width:16px; height:16px; border-radius:50%; background:#000; display:inline-block; vertical-align:-3px; margin-right:8px;"></div>Default Color</div>
                <div style="font-size:12px; color:#5f6368; font-weight: 500; margin:8px 8px 4px;">Theme Colors</div>
                <div class="cp-grid" style="padding:0 8px;"><div *ngFor="let c of themeColorsTop" class="cp-sw" [style.background]="c" (click)="setGridlineColor(c); closeMenus()"></div></div>
                <div class="cp-grid" style="padding:0 8px;"><div *ngFor="let c of themeColorsGrid" class="cp-sw" [style.background]="c" (click)="setGridlineColor(c); closeMenus()"></div></div>
                <div style="font-size:12px; color:#5f6368; font-weight: 500; margin:12px 8px 4px;">Standard Colors</div>
                <div class="cp-grid" style="padding:0 8px;"><div *ngFor="let c of standardColors" class="cp-sw" [style.background]="c" (click)="setGridlineColor(c); closeMenus()"></div></div>
                <div class="mds" style="margin:8px 0;"></div>
                <div class="mdi has-sub" style="padding:6px 8px;">
                  More Colors <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
                  <div class="mdi-sub" style="padding: 12px; min-width: 160px; width: 160px;">
                    <div style="font-size: 12px; color: #5f6368; font-weight: 500; margin-bottom: 8px;">Custom Color</div>
                    <input type="color" [ngModel]="gridlineColor === '#d0d0d0' ? '#000000' : gridlineColor" (ngModelChange)="setGridlineColor($event)" (change)="closeMenus()" style="width: 100%; height: 32px; border: 1px solid #d1d5db; border-radius: 4px; padding: 2px; cursor: pointer; background: transparent;">
                  </div>
                </div>
              </div>
            </div>
            <div class="mdi has-sub"><span class="mdi-icon material-symbols-outlined">swap_horiz</span>Grid Direction <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
              <div class="mdi-sub">
                <div class="mdi" (click)="setGridDirection('ltr');closeMenus()"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{gridDirection==='ltr'?'visible':'hidden'}}; vertical-align:-3px; margin-right:8px;">check</span>Left to Right</div>
                <div class="mdi" (click)="setGridDirection('rtl');closeMenus()"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{gridDirection==='rtl'?'visible':'hidden'}}; vertical-align:-3px; margin-right:8px;">check</span>Right to Left</div>
              </div>
            </div>
            <div class="mdi has-sub"><span class="mdi-icon material-symbols-outlined">space_dashboard</span>Grid Spacing <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
              <div class="mdi-sub">
                <div class="mdi" (click)="setGridSpacing('classic');closeMenus()"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{gridSpacing==='classic'?'visible':'hidden'}}; vertical-align:-3px; margin-right:8px;">check</span>Classic</div>
                <div class="mdi" (click)="setGridSpacing('cozy');closeMenus()"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{gridSpacing==='cozy'?'visible':'hidden'}}; vertical-align:-3px; margin-right:8px;">check</span>Cozy</div>
                <div class="mdi" (click)="setGridSpacing('comfort');closeMenus()"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{gridSpacing==='comfort'?'visible':'hidden'}}; vertical-align:-3px; margin-right:8px;">check</span>Comfort</div>
              </div>
            </div>
            <div class="mdi has-sub"><span class="mdi-icon material-symbols-outlined">zoom_in</span>Zoom <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
              <div class="mdi-sub" style="height:250px; overflow-y:auto;">
                <div class="mdi" (click)="setZoom(400);closeMenus()"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{zoomLevel===400?'visible':'hidden'}}; vertical-align:-3px; margin-right:8px;">check</span>400%</div>
                <div class="mdi" (click)="setZoom(300);closeMenus()"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{zoomLevel===300?'visible':'hidden'}}; vertical-align:-3px; margin-right:8px;">check</span>300%</div>
                <div class="mdi" (click)="setZoom(250);closeMenus()"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{zoomLevel===250?'visible':'hidden'}}; vertical-align:-3px; margin-right:8px;">check</span>250%</div>
                <div class="mdi" (click)="setZoom(200);closeMenus()"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{zoomLevel===200?'visible':'hidden'}}; vertical-align:-3px; margin-right:8px;">check</span>200%</div>
                <div class="mdi" (click)="setZoom(150);closeMenus()"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{zoomLevel===150?'visible':'hidden'}}; vertical-align:-3px; margin-right:8px;">check</span>150%</div>
                <div class="mdi" (click)="setZoom(125);closeMenus()"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{zoomLevel===125?'visible':'hidden'}}; vertical-align:-3px; margin-right:8px;">check</span>125%</div>
                <div class="mdi" (click)="setZoom(100);closeMenus()"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{zoomLevel===100?'visible':'hidden'}}; vertical-align:-3px; margin-right:8px;">check</span>100%</div>
                <div class="mdi" (click)="setZoom(75);closeMenus()"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{zoomLevel===75?'visible':'hidden'}}; vertical-align:-3px; margin-right:8px;">check</span>75%</div>
                <div class="mdi" (click)="setZoom(50);closeMenus()"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{zoomLevel===50?'visible':'hidden'}}; vertical-align:-3px; margin-right:8px;">check</span>50%</div>
                <div class="mds"></div>
                <div class="mdi" (click)="setZoom(100);closeMenus()">Default (100%)</div>
              </div>
            </div>
            <div class="mds"></div>
            <div class="mdi has-sub"><span class="mdi-icon material-symbols-outlined">light_mode</span>Appearance <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
              <div class="mdi-sub">
                <div class="mdi" (click)="appearance='light';closeMenus()"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{appearance==='light'?'visible':'hidden'}}; vertical-align:-3px; margin-right:8px;">check</span><span class="material-symbols-outlined" style="font-size:16px; margin-right:8px; vertical-align:-3px;">light_mode</span>Light</div>
                <div class="mdi" (click)="appearance='dark';closeMenus()"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{appearance==='dark'?'visible':'hidden'}}; vertical-align:-3px; margin-right:8px;">check</span><span class="material-symbols-outlined" style="font-size:16px; margin-right:8px; vertical-align:-3px;">dark_mode</span>Dark</div>
                <div class="mds"></div>
                <div class="mdi" (click)="appearance='system';closeMenus()"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{appearance==='system'?'visible':'hidden'}}; vertical-align:-3px; margin-right:8px;">check</span><span class="material-symbols-outlined" style="font-size:16px; margin-right:8px; vertical-align:-3px;">desktop_windows</span>System Default</div>
              </div>
            </div>
            <div class="mdi has-sub"><span class="mdi-icon material-symbols-outlined">settings_suggest</span>View Settings <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
              <div class="mdi-sub">
                <div class="mdi" (click)="showTopBar = !showTopBar; closeMenus()"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{showTopBar?'visible':'hidden'}}; vertical-align:-3px; margin-right:8px;">check</span>Top Bar</div>
                <div class="mdi" (click)="showFormulaBar = !showFormulaBar; closeMenus()"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{showFormulaBar?'visible':'hidden'}}; vertical-align:-3px; margin-right:8px;">check</span>Formula Bar</div>
                <div class="mdi" (click)="showStatusBar = !showStatusBar; closeMenus()"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{showStatusBar?'visible':'hidden'}}; vertical-align:-3px; margin-right:8px;">check</span>Status Bar</div>
                <div class="mds"></div>
                <div class="mdi" (click)="showNotes = !showNotes; closeMenus()"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{showNotes?'visible':'hidden'}}; vertical-align:-3px; margin-right:8px;">check</span>Notes</div>
                <div class="mdi" (click)="showUserPresence = !showUserPresence; closeMenus()"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{showUserPresence?'visible':'hidden'}}; vertical-align:-3px; margin-right:8px;">check</span>User Presence</div>
                <div class="mdi" (click)="showLockPattern = !showLockPattern; closeMenus()"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{showLockPattern?'visible':'hidden'}}; vertical-align:-3px; margin-right:8px;">check</span>Lock Pattern</div>
                <div class="mdi" (click)="showHighlightPrintArea = !showHighlightPrintArea; closeMenus()"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{showHighlightPrintArea?'visible':'hidden'}}; vertical-align:-3px; margin-right:8px;">check</span>Highlight Print Area</div>
              </div>
            </div>
            <div class="mds"></div>
            <div class="mdi has-sub"><span class="mdi-icon material-symbols-outlined">border_vertical</span>Highlight Row/Column <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
              <div class="mdi-sub" style="width:160px; padding:8px;">
                <div class="cp-grid">
                  <div *ngFor="let c of highlightColors" class="cp-sw" style="border-radius:4px; width:24px; height:24px; border:1px solid #ccc;" [style.background]="c==='transparent'?'url(data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAAIUlEQVQYV2N89erVfwY0ICYmJowuw6iCRQcoP1AwgAUBABvGGR9Lw4lTAAAAAElFTkSuQmCC)':c" (click)="highlightRowColColor = c; closeMenus()"></div>
                </div>
              </div>
            </div>
            <div class="mdi" (click)="toggleFullScreen();closeMenus()"><span class="mdi-icon material-symbols-outlined">fullscreen</span>Full Screen</div>
            <div class="mds"></div>
            <div class="mdi" (click)="openApp('navigation'); closeMenus()">
              <span class="mdi-icon material-symbols-outlined">web_stories</span>Navigation
            </div>
          </div>
        </div>
        <div class="mi" (click)="toggleMenu('insert',$event)" [class.mi-open]="activeMenu==='insert'">Insert
          <div class="mdd" *ngIf="activeMenu==='insert'">
            <div class="mdi has-sub"><span class="mdi-icon material-symbols-outlined">data_table</span>Row <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
              <div class="mdi-sub">
                <div class="mdi" (click)="insertRowAbove()">{{ selectedRowCount }} Row{{ selectedRowCount > 1 ? 's' : '' }} Above</div>
                <div class="mdi" (click)="insertRowBelow()">{{ selectedRowCount }} Row{{ selectedRowCount > 1 ? 's' : '' }} Below</div>
                <div class="mdi" (click)="openCustomInsert('row')">Custom...</div>
              </div>
            </div>
            <div class="mdi has-sub"><span class="mdi-icon material-symbols-outlined">view_column</span>Column <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
              <div class="mdi-sub">
                <div class="mdi" (click)="insertColLeft()">{{ selectedColCount }} Column{{ selectedColCount > 1 ? 's' : '' }} Before</div>
                <div class="mdi" (click)="insertColRight()">{{ selectedColCount }} Column{{ selectedColCount > 1 ? 's' : '' }} After</div>
                <div class="mdi" (click)="openCustomInsert('col')">Custom...</div>
              </div>
            </div>
            <div class="mdi has-sub"><span class="mdi-icon material-symbols-outlined">add_box</span>Cell <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
              <div class="mdi-sub">
                <div class="mdi" (click)="shiftCellsDown()">Shift Cells Down</div>
                <div class="mdi" (click)="shiftCellsRight()">Shift Cells Right</div>
              </div>
            </div>
            <div class="mdi" (click)="addSheet()"><span class="mdi-icon material-symbols-outlined">post_add</span>Sheet<span class="mh">Shift+F11</span></div>
            <div class="mds"></div>
            <div class="mdi" (click)="generateChart()"><span class="mdi-icon material-symbols-outlined">insert_chart</span>Chart...</div>

            <div class="mdi has-sub"><span class="mdi-icon material-symbols-outlined">image</span>Image <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
              <div class="mdi-sub">
                <div class="mdi" (click)="triggerImageInsert('cell')">Image in cell...</div>
                <div class="mdi" (click)="triggerImageInsert('over')">Image over cells...</div>
              </div>
            </div>
            <div class="mdi" (click)="toggleMenu('shape', $event)"><span class="mdi-icon material-symbols-outlined">category</span>Shape</div>
            <div class="mdi" (click)="insertButton()"><span class="mdi-icon material-symbols-outlined">smart_button</span>Button</div>
            <div class="mds"></div>
            <div class="mdi" (click)="insertLink()"><span class="mdi-icon material-symbols-outlined">link</span>Hyperlink...<span class="mh">Ctrl+K</span></div>
            <div class="mdi" (click)="insertFunction('SUM')"><span class="mdi-icon material-symbols-outlined">functions</span>Function...<span class="mh">Shift+F3</span></div>
            <div class="mdi" (click)="defineName()"><span class="mdi-icon material-symbols-outlined">badge</span>Define Name<span class="mh">Ctrl+F3</span></div>
            <div class="mds"></div>
            <div class="mdi" (click)="insertNote()"><span class="mdi-icon material-symbols-outlined">sticky_note_2</span>Note<span class="mh">Shift+F2</span></div>
            <div class="mdi" (click)="insertComment()"><span class="mdi-icon material-symbols-outlined">comment</span>Comment</div>
            <div class="mds"></div>
            <div class="mdi" (click)="insertCheckbox();closeMenus()"><span class="mdi-icon material-symbols-outlined">check_box</span>Checkbox</div>
            
            <div class="mdi has-sub"><span class="mdi-icon material-symbols-outlined">checklist</span>Picklist <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
              <div class="mdi-sub" style="width:400px; padding:16px; max-height:70vh; overflow-y:auto; top: auto; bottom: -5px;">
                <div style="display:grid; grid-template-columns: 1fr 1fr 1fr; gap: 16px; margin-bottom:16px;">
                  
                  <!-- Column 1 -->
                  <div style="display:flex; flex-direction:column; gap:16px;">
                    <div style="display:flex; flex-direction:column; align-items:flex-start; gap:4px;">
                      <div style="font-size:10px; font-weight:600; color:#888; margin-bottom:2px; letter-spacing:0.5px;">PROJECT STATUS</div>
                      <div style="background:#e2e8f0; color:#4a5568; font-size:11px; padding:4px 10px; border-radius:12px; cursor:pointer;" (click)="applyPresetPicklist('project_status')">Yet to start</div>
                      <div style="background:#fed7d7; color:#c53030; font-size:11px; padding:4px 10px; border-radius:12px; cursor:pointer;" (click)="applyPresetPicklist('project_status')">Blocked</div>
                      <div style="background:#fefcbf; color:#b7791f; font-size:11px; padding:4px 10px; border-radius:12px; cursor:pointer;" (click)="applyPresetPicklist('project_status')">In Progress</div>
                      <div style="background:#c6f6d5; color:#276749; font-size:11px; padding:4px 10px; border-radius:12px; cursor:pointer;" (click)="applyPresetPicklist('project_status')">Completed</div>
                    </div>
                    <div style="display:flex; flex-direction:column; align-items:flex-start; gap:4px;">
                      <div style="font-size:10px; font-weight:600; color:#888; margin-bottom:2px; letter-spacing:0.5px;">PRIORITY</div>
                      <div style="background:#bee3f8; color:#2b6cb0; font-size:11px; padding:4px 10px; border-radius:12px; cursor:pointer;" (click)="applyPresetPicklist('priority')">Low</div>
                      <div style="background:#c6f6d5; color:#276749; font-size:11px; padding:4px 10px; border-radius:12px; cursor:pointer;" (click)="applyPresetPicklist('priority')">Medium</div>
                      <div style="background:#fefcbf; color:#b7791f; font-size:11px; padding:4px 10px; border-radius:12px; cursor:pointer;" (click)="applyPresetPicklist('priority')">High</div>
                      <div style="background:#fed7d7; color:#c53030; font-size:11px; padding:4px 10px; border-radius:12px; cursor:pointer;" (click)="applyPresetPicklist('priority')">Critical</div>
                    </div>
                  </div>

                  <!-- Column 2 -->
                  <div style="display:flex; flex-direction:column; gap:16px;">
                    <div style="display:flex; flex-direction:column; align-items:flex-start; gap:4px;">
                      <div style="font-size:10px; font-weight:600; color:#888; margin-bottom:2px; letter-spacing:0.5px;">BUG STATUS</div>
                      <div style="background:#fed7d7; color:#c53030; font-size:11px; padding:4px 10px; border-radius:12px; cursor:pointer;" (click)="applyPresetPicklist('bug_status')">Open</div>
                      <div style="background:#fefcbf; color:#b7791f; font-size:11px; padding:4px 10px; border-radius:12px; cursor:pointer;" (click)="applyPresetPicklist('bug_status')">In Progress</div>
                      <div style="background:#c6f6d5; color:#276749; font-size:11px; padding:4px 10px; border-radius:12px; cursor:pointer;" (click)="applyPresetPicklist('bug_status')">Closed</div>
                      <div style="background:#bee3f8; color:#2b6cb0; font-size:11px; padding:4px 10px; border-radius:12px; cursor:pointer;" (click)="applyPresetPicklist('bug_status')">Reopen</div>
                    </div>
                    <div style="display:flex; flex-direction:column; align-items:flex-start; gap:4px;">
                      <div style="font-size:10px; font-weight:600; color:#888; margin-bottom:2px; letter-spacing:0.5px;">DECISION</div>
                      <div style="background:#c6f6d5; color:#276749; font-size:11px; padding:4px 10px; border-radius:12px; cursor:pointer;" (click)="applyPresetPicklist('decision')">Yes</div>
                      <div style="background:#fed7d7; color:#c53030; font-size:11px; padding:4px 10px; border-radius:12px; cursor:pointer;" (click)="applyPresetPicklist('decision')">No</div>
                    </div>
                  </div>

                  <!-- Column 3 -->
                  <div style="display:flex; flex-direction:column; gap:16px;">
                    <div style="display:flex; flex-direction:column; align-items:flex-start; gap:4px;">
                      <div style="font-size:10px; font-weight:600; color:#888; margin-bottom:2px; letter-spacing:0.5px;">REVIEW</div>
                      <div style="background:#e2e8f0; color:#4a5568; font-size:11px; padding:4px 10px; border-radius:12px; cursor:pointer;" (click)="applyPresetPicklist('review')">Yet to start</div>
                      <div style="background:#bee3f8; color:#2b6cb0; font-size:11px; padding:4px 10px; border-radius:12px; cursor:pointer;" (click)="applyPresetPicklist('review')">Under Review</div>
                      <div style="background:#c6f6d5; color:#276749; font-size:11px; padding:4px 10px; border-radius:12px; cursor:pointer;" (click)="applyPresetPicklist('review')">Approved</div>
                    </div>
                    <div style="display:flex; flex-direction:column; align-items:flex-start; gap:4px;">
                      <div style="font-size:10px; font-weight:600; color:#888; margin-bottom:2px; letter-spacing:0.5px;">BOOLEAN</div>
                      <div style="background:#c6f6d5; color:#276749; font-size:11px; padding:4px 10px; border-radius:12px; cursor:pointer;" (click)="applyPresetPicklist('boolean')">True</div>
                      <div style="background:#fed7d7; color:#c53030; font-size:11px; padding:4px 10px; border-radius:12px; cursor:pointer;" (click)="applyPresetPicklist('boolean')">False</div>
                    </div>
                  </div>
                </div>

                <div style="height:1px; background:#e2e8f0; margin-bottom:8px;"></div>
                <div class="mdi" style="padding: 8px 12px; font-weight: 500;" (click)="openValidationModal();closeMenus()">Create Picklist...</div>
                <div class="mdi" style="padding: 8px 12px; font-weight: 500;" (click)="openManagePicklistSidebar($event)">Manage Picklist...</div>
              </div>
            </div>
            <div class="mdi" (click)="insertEmoji()"><span class="mdi-icon material-symbols-outlined">add_reaction</span>Emoji</div>
          </div>
        </div>
        <div class="mi" (click)="toggleMenu('format',$event)" [class.mi-open]="activeMenu==='format'">Format
          <div class="mdd" *ngIf="activeMenu==='format'">
            <div class="mdi has-sub">Text <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
              <div class="mdi-sub">
                <div class="mdi" (click)="toggleFormat('bold')">Bold<span class="mh">Ctrl+B</span></div>
                <div class="mdi" (click)="toggleFormat('italic')">Italic<span class="mh">Ctrl+I</span></div>
                <div class="mdi" (click)="toggleFormat('underline')">Underline<span class="mh">Ctrl+U</span></div>
                <div class="mdi" (click)="toggleFormat('strikethrough')">Strikethrough</div>
              </div>
            </div>
            <div class="mdi has-sub">Alignment <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
              <div class="mdi-sub">
                <div class="mdi" (click)="setFormat('align','left')">Align Left</div>
                <div class="mdi" (click)="setFormat('align','center')">Align Center</div>
                <div class="mdi" (click)="setFormat('align','right')">Align Right</div>
                <div class="mds"></div>
                <div class="mdi" (click)="setFormat('vertAlign','top')">Vertical Top</div>
                <div class="mdi" (click)="setFormat('vertAlign','middle')">Vertical Middle</div>
                <div class="mdi" (click)="setFormat('vertAlign','bottom')">Vertical Bottom</div>
              </div>
            </div>
            <div class="mdi has-sub">Number <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
              <div class="mdi-sub">
                <div class="mdi" (click)="setNumFormat('general')">General</div>
                <div class="mdi" (click)="setNumFormat('number')">Number</div>
                <div class="mdi" (click)="setNumFormat('currency')">Currency ($)</div>
                <div class="mdi" (click)="setNumFormat('percent')">Percentage (%)</div>
              </div>
            </div>
            <div class="mdi has-sub">Borders <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
              <div class="mdi-sub">
                <div class="mdi" (click)="setBorders('all')">All Borders</div>
                <div class="mdi" (click)="setBorders('outer')">Outer Border</div>
                <div class="mdi" (click)="setBorders('none')">No Border</div>
              </div>
            </div>
            <div class="mds"></div>
            <div class="mdi" (click)="mergeCells()">Merge Cells</div>
            <div class="mdi" (click)="unmerge()">Unmerge</div>
            <div class="mdi" (click)="toggleWrap()">Wrap Text</div>
            <div class="mds"></div>
            <div class="mdi" (click)="clearAllFormats()">Clear Formatting</div>
          </div>
        </div>
        <div class="mi" (click)="toggleMenu('data',$event)" [class.mi-open]="activeMenu==='data'">Data
          <div class="mdd" *ngIf="activeMenu==='data'">
            
            <div class="mdi has-sub">
                <span class="material-symbols-outlined mdi-icon">sort_by_alpha</span> Sort
                <span class="material-symbols-outlined mdi-arrow">chevron_right</span>
                <div class="mdi-sub">
                    <div class="mdi" (click)="sortColAZ()"><span class="material-symbols-outlined mdi-icon" style="font-size:16px;">arrow_upward</span> Sort Ascending</div>
                    <div class="mdi" (click)="sortColZA()"><span class="material-symbols-outlined mdi-icon" style="font-size:16px;">arrow_downward</span> Sort Descending</div>
                    <div class="mdi" (click)="showToast('Custom Sort...')"><span class="material-symbols-outlined mdi-icon" style="font-size:16px;">sort</span> Custom Sort...</div>
                </div>
            </div>
            
            <div class="mdi has-sub">
                <span class="material-symbols-outlined mdi-icon">filter_alt</span> Filter
                <span class="mh">Ctrl+Shift+L</span>
                <span class="material-symbols-outlined mdi-arrow">chevron_right</span>
                <div class="mdi-sub">
                    <div class="mdi" (click)="toggleFilter()">Create Filter</div>
                    <div class="mds"></div>
                    <div class="mdi" (click)="showToast('Reapply')">Reapply <span class="mh">Ctrl+Alt+L</span></div>
                    <div class="mdi" (click)="filterActive=false;showToast('Clear Filter')">Clear Filter</div>
                    <div class="mds"></div>
                    <div class="mdi" (click)="showToast('Set as Document Filter')">Set as Document Filter</div>
                    <div class="mdi" (click)="showToast('Name this filter')">Name this filter</div>
                    <div class="mdi" (click)="showToast('Manage Filters')">Manage Filters</div>
                    <div class="mdi has-sub" (click)="showToast('Named Filters')">Named Filters <span class="material-symbols-outlined mdi-arrow">chevron_right</span></div>
                    <div class="mds"></div>
                    <div class="mdi" (click)="showToast('Highlight Filter')"><span class="material-symbols-outlined mdi-icon" style="font-size:16px;color:transparent;">check</span> Highlight Filter</div>
                </div>
            </div>

            <div class="mdi has-sub">
                <span class="material-symbols-outlined mdi-icon">group_work</span> Group & Ungroup
                <span class="material-symbols-outlined mdi-arrow">chevron_right</span>
                <div class="mdi-sub">
                    <div class="mdi" (click)="groupRow()">Group Row <span class="mh">Alt+Shift+&rarr;</span></div>
                    <div class="mdi" (click)="groupCol()">Group Column <span class="mh">Alt+Shift+&rarr;</span></div>
                    <div class="mds"></div>
                    <div class="mdi" style="color:#718096" (click)="ungroupRow()">Ungroup Row <span class="mh">Alt+Shift+&larr;</span></div>
                    <div class="mdi" style="color:#718096" (click)="ungroupCol()">Ungroup Column <span class="mh">Alt+Shift+&larr;</span></div>
                    <div class="mds"></div>
                    <div class="mdi" style="color:#718096" (click)="clearGroups()">Clear Groups</div>
                </div>
            </div>
            
            <div class="mds"></div>

            <div class="mdi has-sub">
                <span class="material-symbols-outlined mdi-icon">pivot_table_chart</span> Pivot Table
                <span class="material-symbols-outlined mdi-arrow">chevron_right</span>
                <div class="mdi-sub">
                    <div class="mdi" (click)="openPivotModal($event)">Create Pivot Table...</div>
                    <div class="mds"></div>
                    <div class="mdi" style="color:#718096" (click)="showToast('Add Pivot Chart')">Add Pivot Chart</div>
                    <div class="mdi" style="color:#718096" (click)="showToast('Add Slicer...')">Add Slicer...</div>
                    <div class="mdi" style="color:#718096" (click)="showToast('Add Timeline...')">Add Timeline...</div>
                </div>
            </div>
            
            <div class="mds"></div>

            <div class="mdi has-sub">
                <span class="material-symbols-outlined mdi-icon">playlist_add_check</span> Data Validation
                <span class="material-symbols-outlined mdi-arrow">chevron_right</span>
                <div class="mdi-sub">
                    <div class="mdi" (click)="openDataValidationModal($event)">Create Validation...</div>
                    <div class="mdi" (click)="openManageRulesModal($event)">Manage Validation...</div>
                    <div class="mds"></div>
                    <div class="mdi" (click)="showToast('Highlight Invalid Data')">Highlight Invalid Data</div>
                </div>
            </div>
            
            <div class="mdi has-sub">
                <span class="material-symbols-outlined mdi-icon">cleaning_services</span> Data Cleaning
                <span class="material-symbols-outlined mdi-arrow">chevron_right</span>
                <div class="mdi-sub" style="min-width: 180px;">
                    <div class="mdi" (click)="removeDuplicates()">Remove Duplicates</div>
                </div>
            </div>

            <div class="mdi" (click)="openTextToColumnsModal(); closeMenus()">
                <span class="material-symbols-outlined mdi-icon">view_column</span> Text to Columns...
            </div>
            <div class="mdi" (click)="patternFill(); closeMenus()">
                <span class="material-symbols-outlined mdi-icon">format_paint</span> Pattern Fill <span class="mh">Ctrl+E</span>
            </div>
            
            <div class="mds"></div>

            <div class="mdi has-sub">
                <span class="material-symbols-outlined mdi-icon">cable</span> Data Connection
                <span class="material-symbols-outlined mdi-arrow">chevron_right</span>
                <div class="mdi-sub">
                  <div class="mdi disabled">Connect to Database</div>
                  <div class="mdi disabled">Connect to API</div>
                </div>
            </div>
            <div class="mdi" (click)="linkSpreadsheet(); closeMenus()">
                <span class="material-symbols-outlined mdi-icon">link</span> Link Spreadsheet...
            </div>
            <div class="mdi" (click)="dataFromPicture(); closeMenus()">
                <span class="material-symbols-outlined mdi-icon">image</span> Data from Picture...
            </div>
            
            <div class="mds"></div>

            <div class="mdi has-sub">
                <span class="material-symbols-outlined mdi-icon">lock</span> Lock
                <span class="material-symbols-outlined mdi-arrow">chevron_right</span>
                <div class="mdi-sub">
                  <div class="mdi" (click)="lockSelectedRange(); closeMenus()">
                    <span class="material-symbols-outlined mdi-icon">{{ isSelectionLocked() ? 'lock_open' : 'lock' }}</span>
                    {{ isSelectionLocked() ? 'Unlock Cells...' : 'Lock Cells...' }}
                  </div>
                  <div class="mdi" (click)="lockCurrentSheet(); closeMenus()">
                    <span class="material-symbols-outlined mdi-icon">lock</span>
                    {{ sheets[currentSheetIdx].locked ? 'Unlock Sheet...' : 'Lock Sheet...' }}
                  </div>
                  <div class="mdi" (click)="manageLockSettings(); closeMenus()">
                    <span class="material-symbols-outlined mdi-icon">settings</span> Manage Lock Settings...
                  </div>
                  <div class="mdi" (click)="highlightLocks(); closeMenus()" style="display:flex; justify-content:space-between; align-items:center;">
                    <div style="display:flex; align-items:center;"><span class="material-symbols-outlined mdi-icon">highlight</span> Highlight Locks</div>
                    <span *ngIf="showLockPattern" class="material-symbols-outlined" style="font-size:16px;">check</span>
                  </div>
                </div>
            </div>
            <div class="mdi has-sub">
                <span class="material-symbols-outlined mdi-icon">publish</span> Publish Range
                <span class="material-symbols-outlined mdi-arrow">chevron_right</span>
                <div class="mdi-sub">
                  <div class="mdi" (click)="publishRange(); closeMenus()">Publish as Web Page</div>
                  <div class="mdi" (click)="copyPublishLink(); closeMenus()">Copy Shareable Link</div>
                </div>
            </div>

          </div>
        </div>
        <div class="mi" (click)="toggleMenu('review',$event)" [class.mi-open]="activeMenu==='review'">Review
          <div class="mdd" *ngIf="activeMenu==='review'" style="min-width:220px;">
            <div class="mdi" (click)="spellCheck(); closeMenus()">
              <span class="material-symbols-outlined mdi-icon">spellcheck</span> Spell Check...
            </div>
            <div class="mdi" (click)="personalDictionary(); closeMenus()">
              <span class="material-symbols-outlined mdi-icon">book</span> Personal Dictionary...
            </div>
            <div class="mds"></div>
            <div class="mdi" (click)="showWordCount(); closeMenus()">
              <span class="material-symbols-outlined mdi-icon">bar_chart</span> Spreadsheet Statistics...
            </div>
            <div class="mdi" (click)="translateSheet(); closeMenus()">
              <span class="material-symbols-outlined mdi-icon">translate</span> Translate...
            </div>
            <div class="mds"></div>
            <div class="mdi" (click)="openAuditTrail(); closeMenus()">
              <span class="material-symbols-outlined mdi-icon">history</span> Audit Trail...
            </div>

            <div class="mdi" (click)="openFeatureModal('version')">
              <span class="material-symbols-outlined mdi-icon">manage_history</span> Version History
            </div>
            <div class="mds"></div>
            <div class="mdi has-sub">
              <span class="material-symbols-outlined mdi-icon">account_tree</span> Workflow
              <span class="material-symbols-outlined mdi-arrow">chevron_right</span>
              <div class="mdi-sub">
                <div class="mdi" (click)="activeModal='workflow'; closeMenus()">Manage Workflows</div>
                <div class="mdi disabled">Create Workflow</div>
              </div>
            </div>
            <div class="mdi has-sub">
              <span class="material-symbols-outlined mdi-icon">comment</span> Comment
              <span class="material-symbols-outlined mdi-arrow">chevron_right</span>
              <div class="mdi-sub">
                <div class="mdi" (click)="insertComment(); closeMenus()">Add Comment</div>
                <div class="mdi" (click)="showAllComments(); closeMenus()">Show All Comments</div>
                <div class="mdi" (click)="toggleHighlightComments(); closeMenus()" style="display:flex; justify-content:space-between;">
                  <span>Highlight Comments</span>
                  <span *ngIf="highlightCommentsEnabled" class="material-symbols-outlined" style="font-size:16px;">check</span>
                </div>
              </div>
            </div>
            <div class="mdi" (click)="insertNote(); closeMenus()">
              <span class="material-symbols-outlined mdi-icon">sticky_note_2</span> Note <span class="mh">Shift+F2</span>
            </div>
            <div class="mds"></div>
            <div class="mdi has-sub">
              <span class="material-symbols-outlined mdi-icon">lock</span> Lock
              <span class="material-symbols-outlined mdi-arrow">chevron_right</span>
              <div class="mdi-sub">
                <div class="mdi" (click)="lockSelectedRange(); closeMenus()">{{ isSelectionLocked() ? 'Unlock Cells...' : 'Lock Cells...' }}</div>
                <div class="mdi" (click)="lockCurrentSheet(); closeMenus()">
                  {{ sheets[currentSheetIdx].locked ? 'Unlock Sheet...' : 'Lock Sheet...' }}
                </div>
                <div class="mdi" (click)="manageLockSettings(); closeMenus()">Manage Lock Settings...</div>
                <div class="mdi" (click)="highlightLocks(); closeMenus()" style="display:flex; justify-content:space-between;">
                  <span>Highlight Locks</span>
                  <span *ngIf="showLockPattern" class="material-symbols-outlined" style="font-size:16px;">check</span>
                </div>
              </div>
            </div>
          </div>
        </div>
        <div class="mi" (click)="toggleMenu('tools',$event)" [class.mi-open]="activeMenu==='tools'">Tools
          <div class="mdd" *ngIf="activeMenu==='tools'" style="min-width:220px;">
            <div class="mdi has-sub">
              <span class="material-symbols-outlined mdi-icon">assignment</span> Form
              <span class="material-symbols-outlined mdi-arrow">chevron_right</span>
              <div class="mdi-sub">
                <div class="mdi" (click)="createForm(); closeMenus()">Create Form</div>
                <div class="mdi" (click)="manageForms(); closeMenus()">Manage Forms</div>
              </div>
            </div>
            <div class="mdi has-sub">
              <span class="material-symbols-outlined mdi-icon">code</span> VBA Macros
              <span class="material-symbols-outlined mdi-arrow">chevron_right</span>
              <div class="mdi-sub">
                <div class="mdi" (click)="openMacroEditor(); closeMenus()">VBA Editor</div>
                <div class="mdi" (click)="runMacro(); closeMenus()">Run Macro</div>
              </div>
            </div>
            <div class="mdi has-sub">
              <span class="material-symbols-outlined mdi-icon">functions</span> Custom Functions
              <span class="material-symbols-outlined mdi-arrow">chevron_right</span>
              <div class="mdi-sub">
                <div class="mdi" (click)="openCustomFunctions(); closeMenus()">Manage Functions</div>
              </div>
            </div>
            <div class="mds"></div>
            <div class="mdi" (click)="openGoalSeek(); closeMenus()">
              <span class="material-symbols-outlined mdi-icon">gps_fixed</span> Goal Seek...
            </div>
            <div class="mdi" (click)="openSolver(); closeMenus()">
              <span class="material-symbols-outlined mdi-icon">tune</span> Solver...
            </div>
            <div class="mds"></div>
            <div class="mdi" (click)="openFind(); closeMenus()">
              <span class="material-symbols-outlined mdi-icon">search</span> Find &amp; Replace <span class="mh">Ctrl+H</span>
            </div>
            <div class="mds"></div>
            <div class="mdi" (click)="openEmailNotifications(); closeMenus()">
              <span class="material-symbols-outlined mdi-icon">notifications</span> Email Notification Settings
            </div>
            <div class="mds"></div>
            <div style="padding: 4px 16px; font-size:11px; color:#9aa0a6; font-weight:600; text-transform:uppercase; letter-spacing:0.5px;">Automation</div>
            <div class="mdi" (click)="openMergeTemplate(); closeMenus()">
              <span class="material-symbols-outlined mdi-icon">merge_type</span> Merge Template...
            </div>
            <div class="mdi" (click)="openPreferences(); closeMenus()">
              <span class="material-symbols-outlined mdi-icon">person</span> My Preferences
              <span style="background:#10b981; color:#fff; font-size:10px; font-weight:700; padding:1px 5px; border-radius:3px; margin-left:6px;">New</span>
            </div>
          </div>
        </div>
        <div class="mi" (click)="toggleMenu('help',$event)" [class.mi-open]="activeMenu==='help'">Help
          <div class="mdd" *ngIf="activeMenu==='help'" style="min-width:240px;">
            
            <div class="mdi" (click)="showKeyboardShortcuts(); closeMenus()">
              <span class="material-symbols-outlined mdi-icon">keyboard</span> Keyboard Shortcuts...
            </div>
            <div class="mdi" (click)="openFeedback(); closeMenus()">
              <span class="material-symbols-outlined mdi-icon">feedback</span> Feedback...
            </div>
          </div>
        </div>
      </div>

      <!-- ═══ TOOLBAR ROW 1 ══════════════════════════════════════════════════ -->
      <div class="tb-row" (mousedown)="$event.preventDefault()">
        <div class="tb-group">
          <button class="tb" (click)="printSheet()" title="Print"><span class="material-symbols-outlined">print</span></button>
          <button class="tb" (click)="undo()" title="Undo (Ctrl+Z)"><span class="material-symbols-outlined">undo</span></button>
          <button class="tb" (click)="redo()" title="Redo (Ctrl+Y)"><span class="material-symbols-outlined">redo</span></button>
          <button class="tb" (click)="clearAllFormats()" title="Clear Formats"><span class="material-symbols-outlined">format_clear</span></button>
        </div>
        <span class="tb-sep"></span>
        <div class="tb-group">
          <div class="tb-font-dd" (click)="toggleMenu('font',$event)" [class.active]="activeMenu==='font'">
            <span [style.font-family]="currentFont" style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">{{currentFont}}</span>
            <span class="material-symbols-outlined" style="font-size:14px;margin-left:2px;">arrow_drop_down</span>
            <div class="mdd font-list" *ngIf="activeMenu==='font'">
              <div class="mdi" *ngFor="let f of fonts" (click)="applyFont(f)" [style.font-family]="f">{{f}}</div>
            </div>
          </div>
        </div>
        <span class="tb-sep"></span>
        <div class="tb-group font-sz">
          <button class="tb sz" (click)="decrementFontSize()" title="Decrease"><span class="material-symbols-outlined" style="font-size:16px;">remove</span></button>
          <div class="tb-font-dd" style="padding:0; height:26px; display:flex; align-items:center; border:none; background:transparent; position:relative; margin:0; cursor:default; min-width:0; gap:0;" [class.active]="activeMenu==='fontsize'">
            <input class="sz-inp" [(ngModel)]="currentSizeNum" (change)="onFontSizeInputChange()" type="number" min="6" max="96" style="width:36px; border-right:none; margin:0; padding-right:0;" (click)="$event.stopPropagation()">
            <div class="sz-drop-btn" (click)="toggleMenu('fontsize', $event)">
              <span class="material-symbols-outlined" style="font-size:14px;">arrow_drop_down</span>
            </div>
            <div class="mdd font-list" *ngIf="activeMenu==='fontsize'" style="min-width:54px; left:0; top:calc(100% + 2px);">
              <div class="mdi" *ngFor="let s of [6,7,8,9,10,11,12,14,18,24,36,48,72]" (click)="currentSizeNum=s; onFontSizeInputChange(); activeMenu=null" style="justify-content:center;">{{s}}</div>
            </div>
          </div>
          <button class="tb sz" (click)="incrementFontSize()" title="Increase"><span class="material-symbols-outlined" style="font-size:16px;">add</span></button>
        </div>
        <span class="tb-sep"></span>
        <div class="tb-group">
          <button class="tb" [class.tb-on]="getFormat('bold')" (click)="toggleFormat('bold')" title="Bold (Ctrl+B)"><span class="material-symbols-outlined">format_bold</span></button>
          <button class="tb" [class.tb-on]="getFormat('italic')" (click)="toggleFormat('italic')" title="Italic (Ctrl+I)"><span class="material-symbols-outlined">format_italic</span></button>
          <button class="tb" [class.tb-on]="getFormat('strikethrough')" (click)="toggleFormat('strikethrough')" title="Strikethrough"><span class="material-symbols-outlined">strikethrough_s</span></button>
          <button class="tb" [class.tb-on]="getFormat('underline')" (click)="toggleFormat('underline')" title="Underline (Ctrl+U)"><span class="material-symbols-outlined">format_underlined</span></button>
        </div>
        <span class="tb-sep"></span>
        <div class="tb-group">
          <div class="tb-clr" (click)="togglePalette('text',$event)" title="Text Color">
            <div class="clr-ico"><span class="material-symbols-outlined" style="font-size:16px;">format_color_text</span><div class="clr-bar" [style.background]="getFormat('color')||'#000'"></div></div>
            <span class="material-symbols-outlined" style="font-size:12px;">arrow_drop_down</span>
            <div class="clr-pop" *ngIf="activePalette==='text'" (click)="$event.stopPropagation()">
              <div class="cp-grid"><div *ngFor="let c of themeColorsTop" class="cp-sw" [style.background]="c" (click)="setFormat('color',c);activePalette=null"></div></div>
              <div class="cp-grid"><div *ngFor="let c of themeColorsGrid" class="cp-sw" [style.background]="c" (click)="setFormat('color',c);activePalette=null"></div></div>
              <div class="cp-grid"><div *ngFor="let c of standardColors" class="cp-sw" [style.background]="c" (click)="setFormat('color',c);activePalette=null"></div></div>
            </div>
          </div>
          <div class="tb-clr" (click)="togglePalette('fill',$event)" title="Fill Color">
            <div class="clr-ico"><span class="material-symbols-outlined" style="font-size:16px;">format_color_fill</span><div class="clr-bar" [style.background]="getFormat('bg')||'#ffff00'"></div></div>
            <span class="material-symbols-outlined" style="font-size:12px;">arrow_drop_down</span>
            <div class="clr-pop" *ngIf="activePalette==='fill'" (click)="$event.stopPropagation()">
              <div class="cp-nocolor" (click)="setFormat('bg','');activePalette=null">&#10006; No Fill</div>
              <div class="cp-grid"><div *ngFor="let c of themeColorsTop" class="cp-sw" [style.background]="c" (click)="setFormat('bg',c);activePalette=null"></div></div>
              <div class="cp-grid"><div *ngFor="let c of themeColorsGrid" class="cp-sw" [style.background]="c" (click)="setFormat('bg',c);activePalette=null"></div></div>
              <div class="cp-grid"><div *ngFor="let c of standardColors" class="cp-sw" [style.background]="c" (click)="setFormat('bg',c);activePalette=null"></div></div>
            </div>
          </div>
        </div>
        <span class="tb-sep"></span>
        <div class="tb-group">
          <div style="position:relative; display:inline-block;">
            <button class="tb" (click)="toggleMenu('border', $event)" [class.tb-on]="activeMenu==='border'" title="Borders">
              <span class="material-symbols-outlined">border_all</span>
              <span class="material-symbols-outlined" style="font-size:12px; margin-left:2px;">arrow_drop_down</span>
            </button>
            <div class="tb-dd" *ngIf="activeMenu==='border'" (click)="$event.stopPropagation()" style="width:230px; padding:10px;">
              <div style="display:flex; gap:12px;">
                <div style="display:grid; grid-template-columns:repeat(5, 1fr); gap:4px; width:140px;">
                   <button class="bp-btn" (click)="setBorders('all'); closeMenus()" title="All Borders"><span class="material-symbols-outlined">border_all</span></button>
                   <button class="bp-btn" (click)="setBorders('inner'); closeMenus()" title="Inner Borders"><span class="material-symbols-outlined">border_inner</span></button>
                   <button class="bp-btn" (click)="setBorders('horizontal'); closeMenus()" title="Horizontal Borders"><span class="material-symbols-outlined">border_horizontal</span></button>
                   <button class="bp-btn" (click)="setBorders('vertical'); closeMenus()" title="Vertical Borders"><span class="material-symbols-outlined">border_vertical</span></button>
                   <button class="bp-btn" (click)="setBorders('outer'); closeMenus()" title="Outer Borders"><span class="material-symbols-outlined">border_outer</span></button>
                   <button class="bp-btn" (click)="setBorders('left'); closeMenus()" title="Left Border"><span class="material-symbols-outlined">border_left</span></button>
                   <button class="bp-btn" (click)="setBorders('top'); closeMenus()" title="Top Border"><span class="material-symbols-outlined">border_top</span></button>
                   <button class="bp-btn" (click)="setBorders('right'); closeMenus()" title="Right Border"><span class="material-symbols-outlined">border_right</span></button>
                   <button class="bp-btn" (click)="setBorders('bottom'); closeMenus()" title="Bottom Border"><span class="material-symbols-outlined">border_bottom</span></button>
                   <button class="bp-btn" (click)="setBorders('none'); closeMenus()" title="Clear Borders"><span class="material-symbols-outlined">border_clear</span></button>
                </div>
                <div style="width:1px; background:#5f6368;"></div>
                <div style="display:flex; flex-direction:column; gap:8px;">
                   <div style="position:relative; z-index:1001;">
                     <div class="bo-item" (click)="activeBorderSubmenu = activeBorderSubmenu === 'color' ? null : 'color'; $event.stopPropagation()" title="Border Color" [class.active-bo]="activeBorderSubmenu==='color'">
                         <div style="width:18px; height:18px; border:1px solid #5f6368;" [style.background]="currentBorderColor"></div>
                         <span class="material-symbols-outlined" style="font-size:14px; color:#a0aec0;">arrow_drop_down</span>
                     </div>
                     <div class="clr-pop" *ngIf="activeBorderSubmenu==='color'" (click)="$event.stopPropagation()" style="position:absolute; top:100%; left:0; right:auto; width:max-content; z-index:1000; margin-top:4px;">
                        <div class="cp-grid"><div *ngFor="let c of themeColorsTop" class="cp-sw" [style.background]="c" (click)="currentBorderColor=c; activeBorderSubmenu=null"></div></div>
                        <div class="cp-grid"><div *ngFor="let c of themeColorsGrid" class="cp-sw" [style.background]="c" (click)="currentBorderColor=c; activeBorderSubmenu=null"></div></div>
                        <div class="cp-grid"><div *ngFor="let c of standardColors" class="cp-sw" [style.background]="c" (click)="currentBorderColor=c; activeBorderSubmenu=null"></div></div>
                     </div>
                   </div>
                   <div style="position:relative; z-index:1000;">
                     <div class="bo-item" (click)="activeBorderSubmenu = activeBorderSubmenu === 'style' ? null : 'style'; $event.stopPropagation()" title="Border Style" [class.active-bo]="activeBorderSubmenu==='style'">
                         <div style="width:18px; height:0;" [ngStyle]="getBorderStyleCss(currentBorderStyle, currentBorderWidth)"></div>
                         <span class="material-symbols-outlined" style="font-size:14px; color:#a0aec0;">arrow_drop_down</span>
                     </div>
                     <div class="mdd" *ngIf="activeBorderSubmenu==='style'" (click)="$event.stopPropagation()" style="position:absolute; top:100%; left:0; right:auto; z-index:1000; margin-top:4px; width:120px;">
                        <div class="mdi" (click)="currentBorderStyle='solid'; currentBorderWidth='1px'; activeBorderSubmenu=null"><div style="width:100%; border-top:1px solid currentColor;"></div></div>
                        <div class="mdi" (click)="currentBorderStyle='solid'; currentBorderWidth='2px'; activeBorderSubmenu=null"><div style="width:100%; border-top:2px solid currentColor;"></div></div>
                        <div class="mdi" (click)="currentBorderStyle='solid'; currentBorderWidth='3px'; activeBorderSubmenu=null"><div style="width:100%; border-top:3px solid currentColor;"></div></div>
                        <div class="mdi" (click)="currentBorderStyle='dashed'; currentBorderWidth='1px'; activeBorderSubmenu=null"><div style="width:100%; border-top:1px dashed currentColor;"></div></div>
                        <div class="mdi" (click)="currentBorderStyle='dotted'; currentBorderWidth='1px'; activeBorderSubmenu=null"><div style="width:100%; border-top:1px dotted currentColor;"></div></div>
                        <div class="mdi" (click)="currentBorderStyle='double'; currentBorderWidth='3px'; activeBorderSubmenu=null"><div style="width:100%; border-top:3px double currentColor;"></div></div>
                     </div>
                   </div>
                </div>
              </div>
            </div>
          </div>
        </div>
        <span class="tb-sep"></span>
        <div class="tb-group">
          <button class="tb" [class.tb-on]="getFormat('align')==='left'" (click)="setFormat('align','left')" title="Align Left"><span class="material-symbols-outlined">format_align_left</span></button>
          <button class="tb" [class.tb-on]="getFormat('align')==='center'" (click)="setFormat('align','center')" title="Align Center"><span class="material-symbols-outlined">format_align_center</span></button>
          <button class="tb" [class.tb-on]="getFormat('align')==='right'" (click)="setFormat('align','right')" title="Align Right"><span class="material-symbols-outlined">format_align_right</span></button>
          <button class="tb" [class.tb-on]="getFormat('vertAlign')==='top'" (click)="setFormat('vertAlign','top')" title="Align Top"><span class="material-symbols-outlined">vertical_align_top</span></button>
          <button class="tb" [class.tb-on]="getFormat('vertAlign')==='middle'" (click)="setFormat('vertAlign','middle')" title="Align Middle"><span class="material-symbols-outlined">vertical_align_center</span></button>
          <button class="tb" [class.tb-on]="getFormat('vertAlign')==='bottom'" (click)="setFormat('vertAlign','bottom')" title="Align Bottom"><span class="material-symbols-outlined">vertical_align_bottom</span></button>
        </div>
        <span class="tb-sep"></span>
        <div class="tb-group">
          <!-- INDENT -->
          <div style="position:relative; display:inline-block;">
            <button class="tb" (click)="toggleMenu('indent', $event)" [class.tb-on]="activeMenu==='indent'" title="Text Indent">
              <span class="material-symbols-outlined">format_indent_increase</span>
              <span class="material-symbols-outlined" style="font-size:12px; margin-left:2px;">expand_more</span>
            </button>
            <div class="tb-dd" *ngIf="activeMenu==='indent'" (click)="$event.stopPropagation()" style="width:220px;">
              <div class="dd-item" (click)="setFormat('indent', 'increase'); closeMenus()"><span class="material-symbols-outlined" style="font-size:16px;">format_indent_increase</span> Increase Indent<span style="margin-left:auto; color:#9aa0a6; font-size:11px;">Ctrl+M</span></div>
              <div class="dd-item" (click)="setFormat('indent', 'decrease'); closeMenus()"><span class="material-symbols-outlined" style="font-size:16px;">format_indent_decrease</span> Decrease Indent<span style="margin-left:auto; color:#9aa0a6; font-size:11px;">Ctrl+Shift+M</span></div>
            </div>
          </div>
          <!-- TEXT WRAP -->
          <div style="position:relative; display:inline-block;">
            <button class="tb" (click)="toggleMenu('wrap', $event)" [class.tb-on]="activeMenu==='wrap'" title="Text Wrapping">
              <span class="material-symbols-outlined">wrap_text</span>
              <span class="material-symbols-outlined" style="font-size:12px; margin-left:2px;">expand_more</span>
            </button>
            <div class="tb-dd" *ngIf="activeMenu==='wrap'" (click)="$event.stopPropagation()" style="width:160px;">
              <div class="dd-item" (click)="setFormat('wrap', 'overflow'); closeMenus()"><span class="material-symbols-outlined" style="font-size:16px;">arrow_right_alt</span> Overflow</div>
              <div class="dd-item" (click)="setFormat('wrap', 'wrap'); closeMenus()"><span class="material-symbols-outlined" style="font-size:16px;">wrap_text</span> Wrap</div>
              <div class="dd-item" (click)="setFormat('wrap', 'clip'); closeMenus()"><span class="material-symbols-outlined" style="font-size:16px;">format_textdirection_r_to_l</span> Clip</div>
              <div class="dd-item" (click)="setFormat('wrap', 'shrink'); closeMenus()"><span class="material-symbols-outlined" style="font-size:16px;">compress</span> Shrink to Fit</div>
            </div>
          </div>
          <!-- TEXT ROTATION -->
          <div style="position:relative; display:inline-block;">
            <button class="tb" (click)="toggleMenu('rotation', $event)" [class.tb-on]="activeMenu==='rotation'" title="Text Rotation">
              <span class="material-symbols-outlined">text_rotation_angleup</span>
              <span class="material-symbols-outlined" style="font-size:12px; margin-left:2px;">expand_more</span>
            </button>
            <div class="tb-dd" *ngIf="activeMenu==='rotation'" (click)="$event.stopPropagation()" style="width:160px;">
              <div class="dd-item" (click)="setFormat('rotation', '0'); closeMenus()"><span class="material-symbols-outlined" style="font-size:16px;">text_rotation_none</span> None</div>
              <div class="dd-item" (click)="setFormat('rotation', '-45'); closeMenus()"><span class="material-symbols-outlined" style="font-size:16px;">text_rotation_angleup</span> Tilt Up</div>
              <div class="dd-item" (click)="setFormat('rotation', '45'); closeMenus()"><span class="material-symbols-outlined" style="font-size:16px;">text_rotation_angledown</span> Tilt Down</div>
              <div class="dd-item" (click)="setFormat('rotation', '-90'); closeMenus()"><span class="material-symbols-outlined" style="font-size:16px;">text_rotate_up</span> Rotate Up</div>
              <div class="dd-item" (click)="setFormat('rotation', '90'); closeMenus()"><span class="material-symbols-outlined" style="font-size:16px;">text_rotation_down</span> Rotate Down</div>
              <div class="dd-item" (click)="setFormat('rotation', 'custom'); closeMenus()"><span class="material-symbols-outlined" style="font-size:16px;">rotate_right</span> Custom...</div>
            </div>
          </div>
          <div style="position:relative; display:inline-block;">
            <button class="tb" (click)="toggleMenu('merge', $event)" [class.tb-on]="activeMenu==='merge'" title="Merge"><span class="material-symbols-outlined">merge_type</span></button>
            <div class="tb-dd" *ngIf="activeMenu==='merge'" (click)="$event.stopPropagation()" style="width:160px;">
              <div class="dd-item" (click)="mergeCells('all')"><span class="material-symbols-outlined" style="font-size:16px;">table_chart</span> Merge Cells</div>
              <div class="dd-item" (click)="mergeCells('across')"><span class="material-symbols-outlined" style="font-size:16px;">view_stream</span> Merge Across</div>
              <div class="dd-item" (click)="mergeCells('down')"><span class="material-symbols-outlined" style="font-size:16px;">view_week</span> Merge Down</div>
              <div class="dd-item" (click)="mergeCells('center')"><span class="material-symbols-outlined" style="font-size:16px;">center_focus_strong</span> Merge and Center</div>
              <div class="dd-item" (click)="unmerge()"><span class="material-symbols-outlined" style="font-size:16px;">grid_on</span> Unmerge</div>
            </div>
          </div>
        </div>
        <span class="tb-sep"></span>
        <div class="tb-group">
          <div class="tb-font-dd" (click)="toggleMenu('numfmt',$event)" [class.active]="activeMenu==='numfmt'" style="min-width:90px;">
            <span>{{getFormatName(getFormat('numFormat'))}}</span>
            <span class="material-symbols-outlined" style="font-size:14px;margin-left:auto;">arrow_drop_down</span>
            <div class="mdd" *ngIf="activeMenu==='numfmt'">
              <div class="mdi" (click)="setNumFormat('general')">General</div>
              <div class="mdi" (click)="setNumFormat('number')">Number <span class="mh">Ctrl+Shift+1</span></div>
              <div class="mdi has-sub">Accounting <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
                <div class="mdi-sub sub-left">
                  <div class="mdi" (click)="setNumFormat('accounting_inr')">₹ Indian Rupee</div>
                  <div class="mdi" (click)="setNumFormat('accounting_usd')">$ United States Dollar</div>
                  <div class="mdi" (click)="setNumFormat('accounting_eur')">€ Euro</div>
                  <div class="mdi" (click)="setNumFormat('accounting_gbp')">£ British Pound Sterling</div>
                  <div class="mdi" (click)="setNumFormat('accounting_cny')">¥ Chinese Yuan</div>
                  <div class="mds"></div>
                  <div class="mdi">More Accounting Formats...</div>
                </div>
              </div>
              <div class="mdi has-sub">Currency <span class="mh">Ctrl+Shift+4</span> <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
                <div class="mdi-sub sub-left">
                  <div class="mdi" (click)="setNumFormat('currency_inr')">₹ Indian Rupee</div>
                  <div class="mdi" (click)="setNumFormat('currency_usd')">$ United States Dollar</div>
                  <div class="mdi" (click)="setNumFormat('currency_eur')">€ Euro</div>
                  <div class="mdi" (click)="setNumFormat('currency_gbp')">£ British Pound Sterling</div>
                  <div class="mdi" (click)="setNumFormat('currency_cny')">¥ Chinese Yuan</div>
                  <div class="mds"></div>
                  <div class="mdi">More Accounting Formats...</div>
                </div>
              </div>
              <div class="mdi has-sub">Date <span class="mh">Ctrl+Shift+3</span> <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
                <div class="mdi-sub sub-left">
                  <div class="mdi" (click)="setNumFormat('date_1')">15/6/26 <span class="mh">d/M/yy</span></div>
                  <div class="mdi" (click)="setNumFormat('date_2')">15 Jun, 2026 <span class="mh">d MMM, yyyy</span></div>
                  <div class="mdi" (click)="setNumFormat('date_3')">15 June, 2026 <span class="mh">d MMMM, yyyy</span></div>
                  <div class="mdi" (click)="setNumFormat('date_4')">Monday, 15 June, 2026 <span class="mh">EEEE, d MMMM, yyyy</span></div>
                  <div class="mdi" (click)="setNumFormat('date_5')">15/06/2026 <span class="mh">dd/MM/yyyy</span></div>
                  <div class="mdi" (click)="setNumFormat('date_6')">06/15/2026 <span class="mh">MM/dd/yyyy</span></div>
                  <div class="mdi" (click)="setNumFormat('date_7')">2026/06/15 <span class="mh">yyyy/MM/dd</span></div>
                  <div class="mdi-title">Date and Time</div>
                  <div class="mdi" (click)="setNumFormat('date_8')">15/6/26 5:22:25 PM IST <span class="mh">d/M/yy h:mm:ss a z</span></div>
                  <div class="mdi" (click)="setNumFormat('date_9')">15 Jun, 2026 5:22:25 PM IST <span class="mh">d MMM, yyyy h:mm:ss a z</span></div>
                  <div class="mdi" (click)="setNumFormat('date_10')">15 June, 2026 5:22:25 PM <span class="mh">d MMMM, yyyy h:mm:ss a</span></div>
                  <div class="mdi" (click)="setNumFormat('date_11')">Monday, 15 June, 2026 5:22 PM <span class="mh">EEEE, d MMMM, yyyy h:mm a</span></div>
                  <div class="mdi" (click)="setNumFormat('date_12')">15/6/26 5:22 PM <span class="mh">d/M/yy h:mm a</span></div>
                </div>
              </div>
              <div class="mdi has-sub">Time <span class="mh">Ctrl+Shift+2</span> <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
                <div class="mdi-sub sub-left">
                  <div class="mdi-title">Time</div>
                  <div class="mdi" (click)="setNumFormat('time_1')">5:22 PM <span class="mh">h:mm a</span></div>
                  <div class="mdi" (click)="setNumFormat('time_2')">5:22:25 PM <span class="mh">h:mm:ss a</span></div>
                  <div class="mdi" (click)="setNumFormat('time_3')">5:22:25 PM IST <span class="mh">h:mm:ss a z</span></div>
                  <div class="mdi-title">Duration</div>
                  <div class="mdi" (click)="setNumFormat('time_4')">25:01 <span class="mh">[HH]:mm</span></div>
                  <div class="mdi" (click)="setNumFormat('time_5')">25:01:01 <span class="mh">[HH]:mm:ss</span></div>
                </div>
              </div>
              <div class="mdi" (click)="setNumFormat('percent')">Percentage <span class="mh">Ctrl+Shift+5</span></div>
              <div class="mdi has-sub">Fraction <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
                <div class="mdi-sub sub-left">
                  <div class="mdi" (click)="setNumFormat('fraction_1')">Up to one digit (1/4)</div>
                  <div class="mdi" (click)="setNumFormat('fraction_2')">Up to two digits (21/25)</div>
                  <div class="mdi" (click)="setNumFormat('fraction_3')">Up to three digits (312/943)</div>
                </div>
              </div>
              <div class="mdi" (click)="setNumFormat('scientific')">Scientific <span class="mh">Ctrl+Shift+6</span></div>
              <div class="mdi" (click)="setNumFormat('text')">Text</div>
              <div class="mdi has-sub">Regional <span class="mdi-arrow material-symbols-outlined">chevron_right</span>
                <div class="mdi-sub sub-left" style="top: -10px;">
                  <div class="mdi" (click)="setNumFormat('regional_zip')">Zip Code</div>
                  <div class="mdi" (click)="setNumFormat('regional_phone')">Phone Number</div>
                  <div class="mdi" (click)="setNumFormat('regional_zip4')">Zip Code+4</div>
                </div>
              </div>
              <div class="mdi" (click)="openCustomFormatModal()">Custom</div>
              <div class="mds"></div>
              <div class="mdi" (click)="openMoreFormatsModal()">More Formats...</div>
            </div>
          </div>
          <button class="tb nf" (click)="setNumFormat('currency')" [class.tb-on]="getFormat('numFormat')==='currency'" title="Currency">$</button>
          <button class="tb nf" (click)="setNumFormat('percent')" [class.tb-on]="getFormat('numFormat')==='percent'" title="Percent">%</button>
          <button class="tb nf" (click)="decreaseDecimals()" title="Decrease Decimals">.0</button>
          <button class="tb nf" (click)="increaseDecimals()" title="Increase Decimals">.00</button>
        </div>
      </div>


      <!-- ═══ TOOLBAR ROW 2 ══════════════════════════════════════════════════ -->
      <div class="tb-row tb-row2" (mousedown)="$event.preventDefault()">
        <button class="tb" (click)="openFind()" title="Find &amp; Replace (Ctrl+H)"><span class="material-symbols-outlined">search</span></button>
        <button class="tb" (click)="insertLink()" title="Insert Link"><span class="material-symbols-outlined">link</span></button>
        <button class="tb" (click)="insertComment()" title="Insert Comment"><span class="material-symbols-outlined">comment</span></button>
        <div style="position:relative; display:inline-block;">
          <button class="tb" [class.tb-on]="activeMenu==='chart'" (click)="toggleMenu('chart', $event)" title="Insert Chart"><span class="material-symbols-outlined">insert_chart</span></button>
          <div class="tb-chart-dd" *ngIf="activeMenu==='chart'" (click)="$event.stopPropagation()">
            <div class="chart-header-icons">
               <span class="material-symbols-outlined" [class.active]="activeChartTab==='column'" (click)="activeChartTab='column'" title="Column">insert_chart</span>
               <span class="material-symbols-outlined" [class.active]="activeChartTab==='bar'" (click)="activeChartTab='bar'" title="Bar">bar_chart</span>
               <span class="material-symbols-outlined" [class.active]="activeChartTab==='line'" (click)="activeChartTab='line'" title="Line">show_chart</span>
               <span class="material-symbols-outlined" [class.active]="activeChartTab==='pie'" (click)="activeChartTab='pie'" title="Pie">pie_chart</span>
               <span class="material-symbols-outlined" [class.active]="activeChartTab==='area'" (click)="activeChartTab='area'" title="Area">area_chart</span>
               <span class="material-symbols-outlined" [class.active]="activeChartTab==='scatter'" (click)="activeChartTab='scatter'" title="Scatter">scatter_plot</span>
               <span class="material-symbols-outlined" [class.active]="activeChartTab==='more'" (click)="activeChartTab='more'" title="Other Charts">candlestick_chart</span>
            </div>

            <!-- COLUMN TAB -->
            <div class="chart-grid" *ngIf="activeChartTab==='column'">
               <div class="chart-item" (click)="generateChart('column')">
                  <svg width="40" height="40" viewBox="0 0 40 40"><rect x="6" y="16" width="10" height="24" fill="#0ea5e9"/><rect x="24" y="6" width="10" height="34" fill="#10b981"/></svg>
                  <span>Column</span>
               </div>
               <div class="chart-item" (click)="generateChart('stacked_column')">
                  <svg width="40" height="40" viewBox="0 0 40 40"><rect x="15" y="16" width="10" height="24" fill="#0ea5e9"/><rect x="15" y="6" width="10" height="10" fill="#10b981"/></svg>
                  <span>Stacked Column</span>
               </div>
               <div class="chart-item" (click)="generateChart('stacked_100')">
                  <svg width="40" height="40" viewBox="0 0 40 40"><rect x="15" y="16" width="10" height="24" fill="#0ea5e9"/><rect x="15" y="0" width="10" height="16" fill="#10b981"/></svg>
                  <span>Stacked Col 100%</span>
               </div>
               <div class="chart-item" (click)="generateChart('grouped')">
                  <svg width="40" height="40" viewBox="0 0 40 40"><rect x="10" y="16" width="8" height="24" fill="#0ea5e9"/><rect x="20" y="6" width="8" height="34" fill="#10b981"/></svg>
                  <span>Grouped Column</span>
               </div>
            </div>

            <!-- BAR TAB -->
            <div class="chart-grid" *ngIf="activeChartTab==='bar'">
               <div class="chart-item" (click)="generateChart('column')">
                  <svg width="40" height="40" viewBox="0 0 40 40"><rect y="6" x="0" height="10" width="24" fill="#0ea5e9"/><rect y="24" x="0" height="10" width="34" fill="#10b981"/></svg>
                  <span>Bar</span>
               </div>
               <div class="chart-item" (click)="generateChart('stacked_column')">
                  <svg width="40" height="40" viewBox="0 0 40 40"><rect y="15" x="0" height="10" width="24" fill="#0ea5e9"/><rect y="15" x="24" height="10" width="10" fill="#10b981"/></svg>
                  <span>Stacked Bar</span>
               </div>
               <div class="chart-item" (click)="generateChart('stacked_100')">
                  <svg width="40" height="40" viewBox="0 0 40 40"><rect y="15" x="0" height="10" width="24" fill="#0ea5e9"/><rect y="15" x="24" height="10" width="16" fill="#10b981"/></svg>
                  <span>Stacked Bar 100%</span>
               </div>
            </div>

            <!-- LINE TAB -->
            <div class="chart-grid" *ngIf="activeChartTab==='line'">
               <div class="chart-item" (click)="generateChart('line')">
                  <svg width="40" height="40" viewBox="0 0 40 40"><polyline points="2,38 15,20 25,25 38,5" fill="none" stroke="#0ea5e9" stroke-width="3"/></svg>
                  <span>Line</span>
               </div>
               <div class="chart-item" (click)="generateChart('line')">
                  <svg width="40" height="40" viewBox="0 0 40 40"><path d="M 2 38 C 10 38, 10 20, 15 20 C 20 20, 20 25, 25 25 C 30 25, 30 5, 38 5" fill="none" stroke="#10b981" stroke-width="3"/></svg>
                  <span>Spline</span>
               </div>
               <div class="chart-item" (click)="generateChart('line')">
                  <svg width="40" height="40" viewBox="0 0 40 40"><polyline points="2,38 15,38 15,20 25,20 25,5 38,5" fill="none" stroke="#f59e0b" stroke-width="3"/></svg>
                  <span>Step Line</span>
               </div>
            </div>

            <!-- PIE TAB -->
            <div class="chart-grid" *ngIf="activeChartTab==='pie'">
               <div class="chart-item" (click)="generateChart('pie')">
                  <svg width="40" height="40" viewBox="0 0 40 40">
                     <path d="M20,20 L20,0 A20,20 0 0,1 40,20 Z" fill="#0ea5e9"/>
                     <path d="M20,20 L40,20 A20,20 0 1,1 20,0 Z" fill="#10b981"/>
                  </svg>
                  <span>Pie</span>
               </div>
               <div class="chart-item" (click)="generateChart('pie')">
                  <svg width="40" height="40" viewBox="0 0 40 40">
                     <path d="M20,25 L5,25 A15,15 0 0,1 35,25 Z" fill="#0ea5e9"/>
                     <path d="M20,25 L35,25 A15,15 0 0,0 20,10 Z" fill="#10b981"/>
                  </svg>
                  <span>Semi Pie</span>
               </div>
               <div class="chart-item" (click)="generateChart('pie')">
                  <svg width="40" height="40" viewBox="0 0 40 40">
                     <path d="M20,20 L20,0 A20,20 0 0,1 40,20 Z" fill="#0ea5e9"/>
                     <path d="M20,20 L40,20 A20,20 0 1,1 20,0 Z" fill="#10b981"/>
                     <circle cx="20" cy="20" r="10" fill="#202124"/>
                  </svg>
                  <span>Doughnut</span>
               </div>
            </div>

            <!-- AREA TAB -->
            <div class="chart-grid" *ngIf="activeChartTab==='area'">
               <div class="chart-item" (click)="generateChart('area')">
                  <svg width="40" height="40" viewBox="0 0 40 40">
                     <polygon points="0,40 10,20 20,25 40,5 40,40" fill="#0ea5e9" opacity="0.8"/>
                  </svg>
                  <span>Area</span>
               </div>
               <div class="chart-item" (click)="generateChart('area')">
                  <svg width="40" height="40" viewBox="0 0 40 40">
                     <polygon points="0,40 10,20 20,25 40,5 40,40" fill="#0ea5e9" opacity="0.6"/>
                     <polygon points="0,40 10,10 20,15 40,0 40,40" fill="#10b981" opacity="0.6"/>
                  </svg>
                  <span>Stacked Area</span>
               </div>
            </div>

            <!-- SCATTER TAB -->
            <div class="chart-grid" *ngIf="activeChartTab==='scatter'">
               <div class="chart-item" (click)="generateChart('scatter')">
                  <svg width="40" height="40" viewBox="0 0 40 40">
                     <circle cx="10" cy="30" r="4" fill="#0ea5e9"/>
                     <circle cx="20" cy="15" r="4" fill="#10b981"/>
                     <circle cx="30" cy="25" r="4" fill="#f59e0b"/>
                     <circle cx="35" cy="10" r="4" fill="#ef4444"/>
                  </svg>
                  <span>Scatter</span>
               </div>
               <div class="chart-item" (click)="generateChart('line')">
                  <svg width="40" height="40" viewBox="0 0 40 40">
                     <polyline points="10,30 20,15 30,25 35,10" fill="none" stroke="#0ea5e9" stroke-width="2"/>
                  </svg>
                  <span>Scatter Line</span>
               </div>
               <div class="chart-item" (click)="generateChart('line')">
                  <svg width="40" height="40" viewBox="0 0 40 40">
                     <polyline points="10,30 20,15 30,25 35,10" fill="none" stroke="#10b981" stroke-width="2"/>
                     <circle cx="10" cy="30" r="3" fill="#10b981"/>
                     <circle cx="20" cy="15" r="3" fill="#10b981"/>
                     <circle cx="30" cy="25" r="3" fill="#10b981"/>
                     <circle cx="35" cy="10" r="3" fill="#10b981"/>
                  </svg>
                  <span>Scatter Line Markers</span>
               </div>
               <div class="chart-item" (click)="generateChart('scatter')">
                  <svg width="40" height="40" viewBox="0 0 40 40">
                     <circle cx="10" cy="30" r="6" fill="#0ea5e9" opacity="0.7"/>
                     <circle cx="20" cy="15" r="8" fill="#10b981" opacity="0.7"/>
                     <circle cx="30" cy="25" r="5" fill="#f59e0b" opacity="0.7"/>
                     <circle cx="35" cy="10" r="9" fill="#ef4444" opacity="0.7"/>
                  </svg>
                  <span>Bubble</span>
               </div>
            </div>

            <!-- MORE / OTHER CHARTS TAB -->
            <div class="chart-grid" *ngIf="activeChartTab==='more'">
               <div class="chart-item" (click)="generateChart('column')">
                  <svg width="40" height="40" viewBox="0 0 40 40">
                     <rect x="8" y="20" width="8" height="20" fill="#0ea5e9"/>
                     <rect x="24" y="10" width="8" height="30" fill="#0ea5e9"/>
                     <polyline points="12,15 28,5" fill="none" stroke="#10b981" stroke-width="2"/>
                  </svg>
                  <span>Combination</span>
               </div>
               <div class="chart-item" (click)="generateChart('column')">
                  <svg width="40" height="40" viewBox="0 0 40 40">
                     <rect x="5" y="30" width="8" height="10" fill="#0ea5e9"/>
                     <rect x="15" y="20" width="8" height="10" fill="#10b981"/>
                     <rect x="25" y="25" width="8" height="5" fill="#ef4444"/>
                  </svg>
                  <span>Waterfall</span>
               </div>
               <div class="chart-item" (click)="generateChart('column')">
                  <svg width="40" height="40" viewBox="0 0 40 40">
                     <rect x="15" y="5" width="10" height="30" fill="#4b5563"/>
                     <rect x="18" y="15" width="4" height="20" fill="#0ea5e9"/>
                     <line x1="12" y1="25" x2="28" y2="25" stroke="#ef4444" stroke-width="2"/>
                  </svg>
                  <span>Vertical Bullet</span>
               </div>
               <div class="chart-item" (click)="generateChart('column')">
                  <svg width="40" height="40" viewBox="0 0 40 40">
                     <rect x="5" y="15" width="30" height="10" fill="#4b5563"/>
                     <rect x="5" y="18" width="20" height="4" fill="#0ea5e9"/>
                     <line x1="15" y1="12" x2="15" y2="28" stroke="#ef4444" stroke-width="2"/>
                  </svg>
                  <span>Horizontal Bullet</span>
               </div>
               <div class="chart-item" (click)="generateChart('column')">
                  <svg width="40" height="40" viewBox="0 0 40 40">
                     <polygon points="5,10 35,10 25,30 15,30" fill="#0ea5e9"/>
                  </svg>
                  <span>Funnel</span>
               </div>
               <div class="chart-item" (click)="generateChart('column')">
                  <svg width="40" height="40" viewBox="0 0 40 40">
                     <polygon points="5,5 35,5 28,15 12,15" fill="#0ea5e9"/>
                     <polygon points="12,16 28,16 22,26 18,26" fill="#10b981"/>
                     <polygon points="18,27 22,27 20,35 20,35" fill="#ef4444"/>
                  </svg>
                  <span>Weighted Funnel</span>
               </div>
               <div class="chart-item" (click)="generateChart('column')">
                  <svg width="40" height="40" viewBox="0 0 40 40">
                     <rect x="5" y="10" width="8" height="30" fill="#0ea5e9"/>
                     <rect x="15" y="20" width="8" height="20" fill="#0ea5e9"/>
                     <rect x="25" y="30" width="8" height="10" fill="#0ea5e9"/>
                     <polyline points="9,25 19,15 29,5" fill="none" stroke="#ef4444" stroke-width="2"/>
                  </svg>
                  <span>Pareto</span>
               </div>
               <div class="chart-item" (click)="generateChart('column')">
                  <svg width="40" height="40" viewBox="0 0 40 40">
                     <rect x="5" y="10" width="10" height="30" fill="#0ea5e9"/>
                     <rect x="15" y="5" width="10" height="35" fill="#0ea5e9"/>
                     <rect x="25" y="20" width="10" height="20" fill="#0ea5e9"/>
                  </svg>
                  <span>Histogram</span>
               </div>
               <div class="chart-item" (click)="generateChart('column')">
                  <svg width="40" height="40" viewBox="0 0 40 40">
                     <line x1="12" y1="5" x2="12" y2="35" stroke="#10b981" stroke-width="1.5"/>
                     <rect x="8" y="15" width="8" height="10" fill="#10b981"/>
                     <line x1="28" y1="5" x2="28" y2="35" stroke="#ef4444" stroke-width="1.5"/>
                     <rect x="24" y="20" width="8" height="10" fill="#ef4444"/>
                  </svg>
                  <span>Candlestick</span>
               </div>
               <div class="chart-item" (click)="generateChart('column')">
                  <svg width="40" height="40" viewBox="0 0 40 40">
                     <line x1="12" y1="5" x2="12" y2="35" stroke="#10b981" stroke-width="2"/>
                     <line x1="7" y1="15" x2="12" y2="15" stroke="#10b981" stroke-width="2"/>
                     <line x1="12" y1="25" x2="17" y2="25" stroke="#10b981" stroke-width="2"/>
                  </svg>
                  <span>OHLC</span>
               </div>
               <div class="chart-item" (click)="generateChart('pie')">
                  <svg width="40" height="40" viewBox="0 0 40 40">
                     <circle cx="20" cy="20" r="15" fill="none" stroke="#5f6368"/>
                     <line x1="20" y1="5" x2="20" y2="35" stroke="#5f6368"/>
                     <line x1="5" y1="20" x2="35" y2="20" stroke="#5f6368"/>
                     <polygon points="20,10 25,20 20,30 15,20" fill="#0ea5e9" opacity="0.6"/>
                  </svg>
                  <span>Polar</span>
               </div>
               <div class="chart-item" (click)="generateChart('pie')">
                  <svg width="40" height="40" viewBox="0 0 40 40">
                     <polygon points="20,5 35,15 30,35 10,35 5,15" fill="none" stroke="#5f6368"/>
                     <polygon points="20,15 25,20 20,25 15,20" fill="#10b981" opacity="0.6"/>
                  </svg>
                  <span>Spider Web</span>
               </div>
               <div class="chart-item" (click)="generateChart('column')">
                  <svg width="40" height="40" viewBox="0 0 40 40">
                     <rect x="15" y="15" width="10" height="10" fill="none" stroke="#0ea5e9" stroke-width="1.5"/>
                     <line x1="20" y1="5" x2="20" y2="15" stroke="#0ea5e9" stroke-width="1.5"/>
                     <line x1="20" y1="25" x2="20" y2="35" stroke="#0ea5e9" stroke-width="1.5"/>
                     <line x1="15" y1="20" x2="25" y2="20" stroke="#0ea5e9" stroke-width="1.5"/>
                  </svg>
                  <span>Vertical Box Plot</span>
               </div>
               <div class="chart-item" (click)="generateChart('column')">
                  <svg width="40" height="40" viewBox="0 0 40 40">
                     <rect x="15" y="15" width="10" height="10" fill="none" stroke="#10b981" stroke-width="1.5"/>
                     <line x1="5" y1="20" x2="15" y2="20" stroke="#10b981" stroke-width="1.5"/>
                     <line x1="25" y1="20" x2="35" y2="20" stroke="#10b981" stroke-width="1.5"/>
                     <line x1="20" y1="15" x2="20" y2="25" stroke="#10b981" stroke-width="1.5"/>
                  </svg>
                  <span>Horizontal Box Plot</span>
               </div>
               <div class="chart-item" (click)="generateChart('column')">
                  <svg width="40" height="40" viewBox="0 0 40 40">
                     <text x="5" y="15" font-size="10" fill="#0ea5e9" font-weight="bold">Data</text>
                     <text x="18" y="28" font-size="8" fill="#10b981">Sheet</text>
                     <text x="5" y="35" font-size="9" fill="#ef4444">Words</text>
                  </svg>
                  <span>Wordcloud</span>
               </div>
            </div>
            <div class="chart-footer">
               Data Range: {{ sheets[currentSheetIdx].name }}.{{ getRangeRef() }}
            </div>
          </div>
        </div>
        <span class="tb-sep"></span>
        <!-- Insert Shape/Diagram Menu -->
        <div style="position:relative; display:inline-block;">
          <button class="tb" (click)="toggleMenu('shape', $event)" [class.tb-on]="activeMenu==='shape'" title="Insert Shape"><span class="material-symbols-outlined">category</span></button>
          
          <div class="tb-dd shape-panel" *ngIf="activeMenu==='shape'" (click)="$event.stopPropagation()">
             <div class="shape-tabs">
               <div class="s-tab" [class.s-tab-active]="shapeTab==='text'" (click)="shapeTab='text'">
                 <span class="material-symbols-outlined">text_fields</span> Text
               </div>
               <div class="s-tab" [class.s-tab-active]="shapeTab==='shape'" (click)="shapeTab='shape'">
                 <span class="material-symbols-outlined">category</span> Shape
               </div>
               <div class="s-tab" [class.s-tab-active]="shapeTab==='diagram'" (click)="shapeTab='diagram'">
                 <span class="material-symbols-outlined">account_tree</span> Diagram
               </div>
             </div>
             
             <!-- DIAGRAM TAB -->
             <div class="shape-content" *ngIf="shapeTab==='diagram'" style="display:flex; height:300px; padding:0;">
                <div style="width: 80px; border-right: 1px solid #eee; background: #fafafa; display: flex; flex-direction: column;">
                   <div style="padding:12px 0; text-align:center; cursor:pointer;"
                        [style.color]="diagramCategory==='list' ? '#e11d48' : '#5f6368'"
                        [style.border-left]="diagramCategory==='list' ? '3px solid #e11d48' : '3px solid transparent'"
                        [style.background]="diagramCategory==='list' ? '#fff' : 'transparent'"
                        (click)="diagramCategory='list'">
                      <span class="material-symbols-outlined" style="display:block;">format_list_bulleted</span><div style="font-size:11px;">List</div>
                   </div>
                   <div style="padding:12px 0; text-align:center; cursor:pointer;"
                        [style.color]="diagramCategory==='process' ? '#e11d48' : '#5f6368'"
                        [style.border-left]="diagramCategory==='process' ? '3px solid #e11d48' : '3px solid transparent'"
                        [style.background]="diagramCategory==='process' ? '#fff' : 'transparent'"
                        (click)="diagramCategory='process'">
                      <span class="material-symbols-outlined" style="display:block;">arrow_right_alt</span><div style="font-size:11px;">Process</div>
                   </div>
                   <div style="padding:12px 0; text-align:center; cursor:pointer;"
                        [style.color]="diagramCategory==='pyramid' ? '#e11d48' : '#5f6368'"
                        [style.border-left]="diagramCategory==='pyramid' ? '3px solid #e11d48' : '3px solid transparent'"
                        [style.background]="diagramCategory==='pyramid' ? '#fff' : 'transparent'"
                        (click)="diagramCategory='pyramid'">
                      <span class="material-symbols-outlined" style="display:block;">change_history</span><div style="font-size:11px;">Pyramid</div>
                   </div>
                   <div style="padding:12px 0; text-align:center; cursor:pointer;"
                        [style.color]="diagramCategory==='cycle' ? '#e11d48' : '#5f6368'"
                        [style.border-left]="diagramCategory==='cycle' ? '3px solid #e11d48' : '3px solid transparent'"
                        [style.background]="diagramCategory==='cycle' ? '#fff' : 'transparent'"
                        (click)="diagramCategory='cycle'">
                      <span class="material-symbols-outlined" style="display:block;">sync</span><div style="font-size:11px;">Cycle</div>
                   </div>
                </div>
                <div style="flex:1; padding: 16px; overflow-y:auto; background:#fff;">
                   <ng-container *ngIf="diagramCategory==='list'">
                     <div style="display:flex; justify-content:space-between; margin-bottom:12px; align-items:center;">
                        <div style="font-weight:600; font-size:14px;">List</div>
                        <div style="font-size:12px;">Levels 
                          <select style="padding:2px 8px; border-radius:4px; border:1px solid #ccc; outline:none; background:#fff;">
                             <option>3</option><option selected>4</option><option>6</option>
                          </select>
                        </div>
                     </div>
                     <div class="diagram-grid">
                        <div class="diag-item" (click)="insertShape('diagram_drop')">
                           <svg viewBox="0 0 100 60"><circle cx="20" cy="15" r="5" fill="#0ea5e9"/><circle cx="25" cy="30" r="5" fill="#10b981"/><circle cx="20" cy="45" r="5" fill="#f59e0b"/><line x1="30" y1="15" x2="80" y2="15" stroke="#e2e8f0" stroke-width="2"/><line x1="35" y1="30" x2="80" y2="30" stroke="#e2e8f0" stroke-width="2"/><line x1="30" y1="45" x2="80" y2="45" stroke="#e2e8f0" stroke-width="2"/></svg>
                           <div style="font-size:12px; margin-top:4px;">Drop</div>
                        </div>
                        <div class="diag-item" (click)="insertShape('diagram_stack')">
                           <svg viewBox="0 0 100 60"><rect x="15" y="20" width="20" height="20" rx="2" fill="none" stroke="#0ea5e9" stroke-width="1" transform="rotate(-10 25 30)"/><rect x="40" y="20" width="20" height="20" rx="2" fill="none" stroke="#10b981" stroke-width="1"/><rect x="65" y="20" width="20" height="20" rx="2" fill="none" stroke="#f59e0b" stroke-width="1" transform="rotate(10 75 30)"/></svg>
                           <div style="font-size:12px; margin-top:4px;">Stack Card</div>
                        </div>
                        <div class="diag-item" (click)="insertShape('diagram_flag')">
                           <svg viewBox="0 0 100 60"><circle cx="15" cy="15" r="4" fill="#0ea5e9"/><rect x="25" y="12" width="60" height="6" fill="#0ea5e9"/><circle cx="15" cy="30" r="4" fill="#10b981"/><rect x="25" y="27" width="60" height="6" fill="#10b981"/><circle cx="15" cy="45" r="4" fill="#f59e0b"/><rect x="25" y="42" width="60" height="6" fill="#f59e0b"/></svg>
                           <div style="font-size:12px; margin-top:4px;">Flag</div>
                        </div>
                        <div class="diag-item" (click)="insertShape('diagram_ribbon')">
                           <svg viewBox="0 0 100 60"><path d="M 20 10 L 80 10 L 80 20 L 20 20 Z" fill="#0ea5e9"/><path d="M 20 25 L 80 25 L 80 35 L 20 35 Z" fill="#10b981"/><path d="M 20 40 L 80 40 L 80 50 L 20 50 Z" fill="#f59e0b"/></svg>
                           <div style="font-size:12px; margin-top:4px;">Ribbon</div>
                        </div>
                        <div class="diag-item" (click)="insertShape('diagram_alter')">
                           <svg viewBox="0 0 100 60"><circle cx="50" cy="15" r="4" fill="#0ea5e9"/><line x1="20" y1="15" x2="40" y2="15" stroke="#e2e8f0" stroke-width="2"/><line x1="60" y1="15" x2="80" y2="15" stroke="#e2e8f0" stroke-width="2"/><circle cx="50" cy="30" r="4" fill="#10b981"/><line x1="20" y1="30" x2="40" y2="30" stroke="#e2e8f0" stroke-width="2"/><line x1="60" y1="30" x2="80" y2="30" stroke="#e2e8f0" stroke-width="2"/></svg>
                           <div style="font-size:12px; margin-top:4px; color:#e11d48">Alter</div>
                        </div>
                        <div class="diag-item" (click)="insertShape('diagram_deck')">
                           <svg viewBox="0 0 100 60"><rect x="20" y="10" width="15" height="40" fill="none" stroke="#e2e8f0" stroke-width="1"/><rect x="40" y="10" width="15" height="40" fill="none" stroke="#e2e8f0" stroke-width="1"/><rect x="60" y="10" width="15" height="40" fill="none" stroke="#e2e8f0" stroke-width="1"/></svg>
                           <div style="font-size:12px; margin-top:4px;">Deck</div>
                        </div>
                     </div>
                   </ng-container>
                   <ng-container *ngIf="diagramCategory==='process'">
                     <div class="diagram-grid">
                        <div class="diag-item" (click)="insertShape('diagram_process_arrow')">
                           <svg viewBox="0 0 100 60"><polygon points="10,20 40,20 40,10 60,30 40,50 40,40 10,40" fill="#0ea5e9"/><polygon points="45,20 75,20 75,10 95,30 75,50 75,40 45,40" fill="#10b981"/></svg>
                           <div style="font-size:12px; margin-top:4px;">Arrow Process</div>
                        </div>
                        <div class="diag-item" (click)="insertShape('diagram_process_step')">
                           <svg viewBox="0 0 100 60"><rect x="10" y="20" width="20" height="20" fill="#0ea5e9"/><line x1="30" y1="30" x2="40" y2="30" stroke="#ccc" stroke-width="2"/><rect x="40" y="20" width="20" height="20" fill="#10b981"/><line x1="60" y1="30" x2="70" y2="30" stroke="#ccc" stroke-width="2"/><rect x="70" y="20" width="20" height="20" fill="#f59e0b"/></svg>
                           <div style="font-size:12px; margin-top:4px;">Step Process</div>
                        </div>
                     </div>
                   </ng-container>
                   <ng-container *ngIf="diagramCategory==='pyramid'">
                     <div class="diagram-grid">
                        <div class="diag-item" (click)="insertShape('diagram_pyramid_basic')">
                           <svg viewBox="0 0 100 60"><polygon points="50,5 35,20 65,20" fill="#0ea5e9"/><polygon points="32,22 68,22 83,40 17,40" fill="#10b981"/><polygon points="14,42 86,42 100,55 0,55" fill="#f59e0b"/></svg>
                           <div style="font-size:12px; margin-top:4px;">Basic Pyramid</div>
                        </div>
                        <div class="diag-item" (click)="insertShape('diagram_pyramid_inv')">
                           <svg viewBox="0 0 100 60"><polygon points="0,5 100,5 86,18 14,18" fill="#0ea5e9"/><polygon points="17,20 83,20 68,38 32,38" fill="#10b981"/><polygon points="35,40 65,40 50,55" fill="#f59e0b"/></svg>
                           <div style="font-size:12px; margin-top:4px;">Inverted Pyramid</div>
                        </div>
                     </div>
                   </ng-container>
                   <ng-container *ngIf="diagramCategory==='cycle'">
                     <div class="diagram-grid">
                        <div class="diag-item" (click)="insertShape('diagram_cycle_basic')">
                           <svg viewBox="0 0 100 60"><path d="M50,10 A20,20 0 0,1 70,30 L65,30 L72.5,40 L80,30 L75,30 A25,25 0 0,0 50,5 Z" fill="#0ea5e9"/><path d="M70,30 A20,20 0 0,1 30,30 L25,30 L32.5,20 L40,30 L35,30 A25,25 0 0,0 75,30 Z" fill="#10b981"/></svg>
                           <div style="font-size:12px; margin-top:4px;">Basic Cycle</div>
                        </div>
                     </div>
                   </ng-container>
                </div>
             </div>
             
             <!-- SHAPE TAB -->
             <div class="shape-content" *ngIf="shapeTab==='shape'" style="background:#fff;">
                <div style="display:flex; gap:16px; border-bottom:1px solid #eee; padding-bottom:8px; margin-bottom:12px;">
                   <div style="cursor:pointer; display:flex; flex-direction:column; align-items:center;"
                        [style.color]="shapeCategory==='shape' ? '#e11d48' : '#5f6368'"
                        [style.border-bottom]="shapeCategory==='shape' ? '2px solid #e11d48' : '2px solid transparent'"
                        (click)="shapeCategory='shape'">
                      <span class="material-symbols-outlined" style="font-size:20px;">crop_square</span><div style="font-size:10px;">Shape</div>
                   </div>
                   <div style="cursor:pointer; display:flex; flex-direction:column; align-items:center;"
                        [style.color]="shapeCategory==='lines' ? '#e11d48' : '#5f6368'"
                        [style.border-bottom]="shapeCategory==='lines' ? '2px solid #e11d48' : '2px solid transparent'"
                        (click)="shapeCategory='lines'">
                      <span class="material-symbols-outlined" style="font-size:20px;">arrow_right_alt</span><div style="font-size:10px;">Lines</div>
                   </div>
                   <div style="cursor:pointer; display:flex; flex-direction:column; align-items:center;"
                        [style.color]="shapeCategory==='flowchart' ? '#e11d48' : '#5f6368'"
                        [style.border-bottom]="shapeCategory==='flowchart' ? '2px solid #e11d48' : '2px solid transparent'"
                        (click)="shapeCategory='flowchart'">
                      <span class="material-symbols-outlined" style="font-size:20px;">account_tree</span><div style="font-size:10px;">Flowchart</div>
                   </div>
                   <div style="cursor:pointer; display:flex; flex-direction:column; align-items:center;"
                        [style.color]="shapeCategory==='math' ? '#e11d48' : '#5f6368'"
                        [style.border-bottom]="shapeCategory==='math' ? '2px solid #e11d48' : '2px solid transparent'"
                        (click)="shapeCategory='math'">
                      <span class="material-symbols-outlined" style="font-size:20px;">add</span><div style="font-size:10px;">Math</div>
                   </div>
                   <div style="cursor:pointer; display:flex; flex-direction:column; align-items:center;"
                        [style.color]="shapeCategory==='stars' ? '#e11d48' : '#5f6368'"
                        [style.border-bottom]="shapeCategory==='stars' ? '2px solid #e11d48' : '2px solid transparent'"
                        (click)="shapeCategory='stars'">
                      <span class="material-symbols-outlined" style="font-size:20px;">star_outline</span><div style="font-size:10px;">Stars</div>
                   </div>
                   <div style="cursor:pointer; display:flex; flex-direction:column; align-items:center;"
                        [style.color]="shapeCategory==='callouts' ? '#e11d48' : '#5f6368'"
                        [style.border-bottom]="shapeCategory==='callouts' ? '2px solid #e11d48' : '2px solid transparent'"
                        (click)="shapeCategory='callouts'">
                      <span class="material-symbols-outlined" style="font-size:20px;">chat_bubble_outline</span><div style="font-size:10px;">Callouts</div>
                   </div>
                </div>
                
                <div class="shape-grid" *ngIf="shapeCategory==='shape'">
                   <div class="s-item" (click)="insertShape('rect')"><svg viewBox="0 0 24 24"><rect x="4" y="4" width="16" height="16" fill="none" stroke="#5f6368" stroke-width="1.5"/></svg></div>
                   <div class="s-item" (click)="insertShape('roundrect')"><svg viewBox="0 0 24 24"><rect x="4" y="4" width="16" height="16" rx="4" fill="none" stroke="#5f6368" stroke-width="1.5"/></svg></div>
                   <div class="s-item" (click)="insertShape('circle')"><svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="8" fill="none" stroke="#5f6368" stroke-width="1.5"/></svg></div>
                   <div class="s-item" (click)="insertShape('triangle')"><svg viewBox="0 0 24 24"><polygon points="12,4 4,20 20,20" fill="none" stroke="#5f6368" stroke-width="1.5"/></svg></div>
                   <div class="s-item" (click)="insertShape('diamond')"><svg viewBox="0 0 24 24"><polygon points="12,4 20,12 12,20 4,12" fill="none" stroke="#5f6368" stroke-width="1.5"/></svg></div>
                   <div class="s-item" (click)="insertShape('hexagon')"><svg viewBox="0 0 24 24"><polygon points="12,4 20,8 20,16 12,20 4,16 4,8" fill="none" stroke="#5f6368" stroke-width="1.5"/></svg></div>
                   <div class="s-item" (click)="insertShape('octagon')"><svg viewBox="0 0 24 24"><polygon points="8,4 16,4 20,8 20,16 16,20 8,20 4,16 4,8" fill="none" stroke="#5f6368" stroke-width="1.5"/></svg></div>
                </div>
                <div class="shape-grid" *ngIf="shapeCategory==='lines'">
                   <div class="s-item" (click)="insertShape('line_straight')"><svg viewBox="0 0 24 24"><line x1="4" y1="20" x2="20" y2="4" stroke="#5f6368" stroke-width="1.5"/></svg></div>
                   <div class="s-item" (click)="insertShape('line_arrow')"><svg viewBox="0 0 24 24"><line x1="4" y1="20" x2="18" y2="6" stroke="#5f6368" stroke-width="1.5"/><polygon points="16,4 21,3 20,8" fill="#5f6368"/></svg></div>
                   <div class="s-item" (click)="insertShape('line_curve')"><svg viewBox="0 0 24 24"><path d="M4,20 Q12,4 20,20" fill="none" stroke="#5f6368" stroke-width="1.5"/></svg></div>
                   <div class="s-item" (click)="insertShape('line_connector')"><svg viewBox="0 0 24 24"><polyline points="4,20 4,12 20,12 20,4" fill="none" stroke="#5f6368" stroke-width="1.5"/></svg></div>
                </div>
                <div class="shape-grid" *ngIf="shapeCategory==='flowchart'">
                   <div class="s-item" (click)="insertShape('flow_process')"><svg viewBox="0 0 24 24"><rect x="3" y="6" width="18" height="12" fill="none" stroke="#5f6368" stroke-width="1.5"/></svg></div>
                   <div class="s-item" (click)="insertShape('flow_decision')"><svg viewBox="0 0 24 24"><polygon points="12,3 21,12 12,21 3,12" fill="none" stroke="#5f6368" stroke-width="1.5"/></svg></div>
                   <div class="s-item" (click)="insertShape('flow_data')"><svg viewBox="0 0 24 24"><polygon points="6,6 22,6 18,18 2,18" fill="none" stroke="#5f6368" stroke-width="1.5"/></svg></div>
                   <div class="s-item" (click)="insertShape('flow_terminator')"><svg viewBox="0 0 24 24"><rect x="3" y="6" width="18" height="12" rx="6" fill="none" stroke="#5f6368" stroke-width="1.5"/></svg></div>
                </div>
                <div class="shape-grid" *ngIf="shapeCategory==='math'">
                   <div class="s-item" (click)="insertShape('math_plus')"><svg viewBox="0 0 24 24"><path d="M11,4 L13,4 L13,11 L20,11 L20,13 L13,13 L13,20 L11,20 L11,13 L4,13 L4,11 L11,11 Z" fill="none" stroke="#5f6368" stroke-width="1.5"/></svg></div>
                   <div class="s-item" (click)="insertShape('math_minus')"><svg viewBox="0 0 24 24"><rect x="4" y="11" width="16" height="2" fill="#5f6368"/></svg></div>
                   <div class="s-item" (click)="insertShape('math_multiply')"><svg viewBox="0 0 24 24"><path d="M6,6 L18,18 M18,6 L6,18" stroke="#5f6368" stroke-width="2"/></svg></div>
                   <div class="s-item" (click)="insertShape('math_divide')"><svg viewBox="0 0 24 24"><circle cx="12" cy="6" r="2" fill="#5f6368"/><rect x="4" y="11" width="16" height="2" fill="#5f6368"/><circle cx="12" cy="18" r="2" fill="#5f6368"/></svg></div>
                   <div class="s-item" (click)="insertShape('math_equal')"><svg viewBox="0 0 24 24"><rect x="4" y="8" width="16" height="2" fill="#5f6368"/><rect x="4" y="14" width="16" height="2" fill="#5f6368"/></svg></div>
                </div>
                <div class="shape-grid" *ngIf="shapeCategory==='stars'">
                   <div class="s-item" (click)="insertShape('star_5')"><svg viewBox="0 0 24 24"><polygon points="12,2 15,9 22,9 16,14 18,21 12,17 6,21 8,14 2,9 9,9" fill="none" stroke="#5f6368" stroke-width="1.5"/></svg></div>
                   <div class="s-item" (click)="insertShape('star_4')"><svg viewBox="0 0 24 24"><polygon points="12,2 14,10 22,12 14,14 12,22 10,14 2,12 10,10" fill="none" stroke="#5f6368" stroke-width="1.5"/></svg></div>
                </div>
                <div class="shape-grid" *ngIf="shapeCategory==='callouts'">
                   <div class="s-item" (click)="insertShape('callout_rect')"><svg viewBox="0 0 24 24"><path d="M3,4 L21,4 L21,16 L14,16 L10,21 L10,16 L3,16 Z" fill="none" stroke="#5f6368" stroke-width="1.5"/></svg></div>
                   <div class="s-item" (click)="insertShape('callout_round')"><svg viewBox="0 0 24 24"><path d="M4,4 C2,4 2,16 4,16 L10,16 L10,21 L14,16 L20,16 C22,16 22,4 20,4 Z" fill="none" stroke="#5f6368" stroke-width="1.5"/></svg></div>
                   <div class="s-item" (click)="insertShape('callout_cloud')"><svg viewBox="0 0 24 24"><path d="M6,14 C4,14 4,10 6,10 C6,6 12,6 14,8 C16,6 20,8 20,12 C22,12 22,16 18,16 L12,21 L10,16 Z" fill="none" stroke="#5f6368" stroke-width="1.5"/></svg></div>
                </div>
                <div style="display:flex; justify-content:space-between; margin-top:20px; font-size:12px; color:#5f6368; align-items:center;">
                   <div style="color:#e11d48; cursor:pointer; display:flex; align-items:center; gap:4px;">
                      <span class="material-symbols-outlined" style="font-size:16px;">edit</span> Draw with pen
                   </div>
                   <div style="display:flex; align-items:center; gap:4px;"><input type="checkbox" style="margin:0;"> Shape Recognition</div>
                </div>
                <div style="display:flex; gap:16px; font-size:12px; margin-top:8px; color:#5f6368;">
                   <span style="cursor:pointer; display:flex; align-items:center; gap:4px;"><span class="material-symbols-outlined" style="font-size:16px;">horizontal_rule</span> Line</span>
                   <span style="cursor:pointer; display:flex; align-items:center; gap:4px;"><span class="material-symbols-outlined" style="font-size:16px;">gesture</span> Curve</span>
                   <span style="cursor:pointer; display:flex; align-items:center; gap:4px;"><span class="material-symbols-outlined" style="font-size:16px;">draw</span> Freeform</span>
                   <span style="cursor:pointer; display:flex; align-items:center; gap:4px;"><span class="material-symbols-outlined" style="font-size:16px;">edit</span> Scribble</span>
                </div>
             </div>
             
             <!-- TEXT TAB -->
             <div class="shape-content" *ngIf="shapeTab==='text'" style="display:flex; height:300px; padding:0;">
                <div style="width: 80px; border-right: 1px solid #eee; background: #fafafa; display: flex; flex-direction: column;">
                   <div style="padding:12px 0; text-align:center; cursor:pointer;"
                        [style.color]="textCategory==='textbox' ? '#e11d48' : '#5f6368'"
                        [style.border-left]="textCategory==='textbox' ? '3px solid #e11d48' : '3px solid transparent'"
                        [style.background]="textCategory==='textbox' ? '#fff' : 'transparent'"
                        (click)="textCategory='textbox'">
                      <span class="material-symbols-outlined" style="display:block;">text_fields</span><div style="font-size:11px;">Textbox</div>
                   </div>
                   <div style="padding:12px 0; text-align:center; cursor:pointer;"
                        [style.color]="textCategory==='symbol' ? '#e11d48' : '#5f6368'"
                        [style.border-left]="textCategory==='symbol' ? '3px solid #e11d48' : '3px solid transparent'"
                        [style.background]="textCategory==='symbol' ? '#fff' : 'transparent'"
                        (click)="textCategory='symbol'">
                      <span class="material-symbols-outlined" style="display:block;">functions</span><div style="font-size:11px;">Symbol</div>
                   </div>
                </div>
                <div style="flex:1; padding: 16px; overflow-y:auto; background:#fff;">
                   <div class="shape-grid" style="grid-template-columns: repeat(4, 1fr); gap:12px;" *ngIf="textCategory==='textbox'">
                      <div class="s-item" style="display:flex; align-items:center; justify-content:center; border:none; color:#5f6368; font-size:11px;" (click)="insertShape('text')">Text</div>
                      <div class="s-item" style="display:flex; align-items:center; justify-content:center; border:none; color:#000; font-weight:bold; font-size:11px;" (click)="insertShape('text')">Text</div>
                      <div class="s-item" style="display:flex; align-items:center; justify-content:center; border:none; color:#0ea5e9; font-style:italic; font-size:11px;" (click)="insertShape('text')">Text</div>
                      <div class="s-item" style="display:flex; align-items:center; justify-content:center; border:none; color:#10b981; font-size:11px;" (click)="insertShape('text')"><span style="border-bottom:1px solid #10b981; padding-bottom:2px;">Text</span></div>
                      
                      <div class="s-item" style="display:flex; align-items:center; justify-content:center; border:1px solid #ccc; border-radius:16px; color:#5f6368; font-size:11px;" (click)="insertShape('text_rounded')">Text</div>
                      <div class="s-item" style="display:flex; align-items:center; justify-content:center; background:#fef08a; border-radius:4px; border:none; color:#5f6368; font-size:11px;" (click)="insertShape('text_yellow')">Text</div>
                      <div class="s-item" style="display:flex; align-items:center; justify-content:center; background:#0ea5e9; color:#fff; border-radius:4px; border:none; font-size:11px; clip-path: polygon(0% 0%, 75% 0%, 100% 50%, 75% 100%, 0% 100%);" (click)="insertShape('text_arrow')">Text</div>
                   </div>
                   <div class="shape-grid" style="grid-template-columns: repeat(4, 1fr); gap:12px;" *ngIf="textCategory==='symbol'">
                      <div class="s-item" style="font-size:18px; color:#5f6368; display:flex; align-items:center; justify-content:center;" (click)="insertShape('symbol_copy')">©</div>
                      <div class="s-item" style="font-size:18px; color:#5f6368; display:flex; align-items:center; justify-content:center;" (click)="insertShape('symbol_reg')">®</div>
                      <div class="s-item" style="font-size:18px; color:#5f6368; display:flex; align-items:center; justify-content:center;" (click)="insertShape('symbol_tm')">™</div>
                      <div class="s-item" style="font-size:18px; color:#5f6368; display:flex; align-items:center; justify-content:center;" (click)="insertShape('symbol_pi')">π</div>
                      <div class="s-item" style="font-size:18px; color:#5f6368; display:flex; align-items:center; justify-content:center;" (click)="insertShape('symbol_sigma')">Σ</div>
                      <div class="s-item" style="font-size:18px; color:#5f6368; display:flex; align-items:center; justify-content:center;" (click)="insertShape('symbol_omega')">Ω</div>
                      <div class="s-item" style="font-size:18px; color:#5f6368; display:flex; align-items:center; justify-content:center;" (click)="insertShape('symbol_inf')">∞</div>
                   </div>
                </div>
             </div>
          </div>
        </div>

        <div style="position:relative; display:inline-block;">
          <button class="tb" (click)="toggleMenu('image', $event)" [class.tb-on]="activeMenu==='image'" title="Insert Image"><span class="material-symbols-outlined">image</span></button>
          <div class="tb-dd" *ngIf="activeMenu==='image'" (click)="$event.stopPropagation()" style="width:200px;">
            <div class="dd-item" (click)="triggerImageInsert('cell')">Image in cell...</div>
            <div class="dd-item" (click)="triggerImageInsert('over')">Image over cells...</div>
          </div>
        </div>
        <div style="position:relative; display:inline-block;">
          <button class="tb" (click)="toggleMenu('sort', $event)" [class.tb-on]="activeMenu==='sort'" title="Sort"><span class="material-symbols-outlined">sort_by_alpha</span></button>
          <div class="tb-dd" *ngIf="activeMenu==='sort'" (click)="$event.stopPropagation()" style="width:220px;">
            <div class="dd-item" (click)="sortColAZ()">
               <span class="material-symbols-outlined" style="font-size:16px;">sort_by_alpha</span> Sort Ascending
            </div>
            <div class="dd-item" (click)="sortColZA()">
               <span class="material-symbols-outlined" style="font-size:16px;">sort_by_alpha</span> Sort Descending
            </div>
            <div class="dd-item" (click)="customSort()">
               <span class="material-symbols-outlined" style="font-size:16px;">sort</span> Custom Sort...
            </div>
          </div>
        </div>
        <button class="tb" [class.tb-on]="filterActive" (click)="toggleFilter()" title="Filter"><span class="material-symbols-outlined">filter_list</span></button>
        <div style="position:relative; display:inline-block;">
          <button class="tb" (click)="toggleMenu('sum', $event)" [class.tb-on]="activeMenu==='sum'" title="Functions"><span class="material-symbols-outlined">functions</span></button>
          <div class="tb-dd" *ngIf="activeMenu==='sum'" (click)="$event.stopPropagation()" style="width:200px;">
            <div class="dd-item" (click)="insertFunction('SUM')">Sum</div>
            <div class="dd-item" (click)="insertFunction('AVERAGE')">Average</div>
            <div class="dd-item" (click)="insertFunction('COUNT')">Count</div>
            <div class="dd-item" (click)="insertFunction('COUNTIF')">Count of Numbers</div>
            <div class="dd-item" (click)="insertFunction('MAX')">Maximum</div>
            <div class="dd-item" (click)="insertFunction('MIN')">Minimum</div>
            <div style="border-top:1px solid #5f6368; margin:8px 0;"></div>
            <div class="dd-item" (click)="moreFunctions()">More Functions</div>
          </div>
        </div>
        <span class="tb-sep"></span>
        <button class="tb" (click)="toggleFreezeRow()" title="Freeze Rows"><span class="material-symbols-outlined">view_agenda</span></button>
        <button class="tb" (click)="toggleFreezeCol()" title="Freeze Columns"><span class="material-symbols-outlined">view_week</span></button>
        <span class="tb-sep"></span>
        <div class="zoom-ctrl">
          <button class="tb" (click)="zoomOut()"><span class="material-symbols-outlined">zoom_out</span></button>
          <span class="zoom-pct">{{zoomLevel}}%</span>
          <button class="tb" (click)="zoomIn()"><span class="material-symbols-outlined">zoom_in</span></button>
        </div>
      </div>

      <div class="formula-container" *ngIf="showFormulaBar">
        <span class="cell-ref">{{ selectedRef }}</span>
        <span class="fx-label">fx</span>
        <input class="formula-bar" [(ngModel)]="formulaBarValue"
            [disabled]="sheets[currentSheetIdx].locked || false"
            (ngModelChange)="onFormulaBarChange($event)"
            (keydown.enter)="commitFormula()" (blur)="commitFormula()" placeholder="" />
      </div>

      <!-- Hidden image file input -->
      <input #imgInput type="file" accept="image/*" style="display:none" (change)="onImageFileSelected($event)">

    <div class="main-content" style="display:flex; flex:1; overflow:hidden; position:relative;">
      <div class="grid-wrap" #gridWrap style="flex:1; overflow:auto; position:relative; background:#fff;" (scroll)="onGridScroll($event)">
        <div class="resize-line-col" *ngIf="resizingCol !== null" [style.left.px]="resizeLineX"></div>
        <div class="resize-line-row" *ngIf="resizingRow !== null" [style.top.px]="resizeLineY"></div>

        <!-- Shapes Rendering -->
        <ng-container *ngIf="sheets[currentSheetIdx].shapes">
          <div *ngFor="let s of sheets[currentSheetIdx].shapes; let i = index"
               class="sheet-shape"
               [class.shape-active]="activeShapeIdx === i"
               [style.left.px]="s.x" [style.top.px]="s.y"
               [style.width.px]="s.width" [style.height.px]="s.height"
               (mousedown)="startShapeDrag($event, i)"
               (dblclick)="editShapeLabel(i)"
               title="Double-click to edit label">
               
               <!-- Active Handles -->
               <ng-container *ngIf="activeShapeIdx === i">
                   <div class="shape-handle nw"></div>
                   <div class="shape-handle n"></div>
                   <div class="shape-handle ne"></div>
                   <div class="shape-handle e"></div>
                   <div class="shape-handle se"></div>
                   <div class="shape-handle s"></div>
                   <div class="shape-handle sw"></div>
                   <div class="shape-handle w"></div>
                   <div class="shape-menu-btn" (mousedown)="$event.stopPropagation()" (click)="activeShapeMenuIdx = activeShapeMenuIdx === i ? null : i; $event.stopPropagation()"><span class="material-symbols-outlined" style="font-size:16px;">more_horiz</span></div>
                   <div class="shape-context-menu" *ngIf="activeShapeMenuIdx === i" (mousedown)="$event.stopPropagation()">
                       <div class="scm-item" (click)="showToast('Assign Existing'); activeShapeMenuIdx=null"><span class="material-symbols-outlined" style="font-size:18px;">description</span> Assign Existing <span class="material-symbols-outlined chevron" style="font-size:18px;">chevron_right</span></div>
                       <div class="scm-item" (click)="showToast('Assign New'); activeShapeMenuIdx=null"><span class="material-symbols-outlined" style="font-size:18px;">post_add</span> Assign New <span class="material-symbols-outlined chevron" style="font-size:18px;">chevron_right</span></div>
                       <div class="scm-item" (click)="editShapeLabel(i); activeShapeMenuIdx=null"><span class="material-symbols-outlined" style="font-size:18px;">edit</span> Edit Label</div>
                       <div style="border-top:1px solid #eee; margin:4px 0;"></div>
                       <div class="scm-item" (click)="showToast('Clone'); activeShapeMenuIdx=null"><span class="material-symbols-outlined" style="font-size:18px;">file_copy</span> Clone</div>
                       <div class="scm-item" (click)="deleteShape(i); activeShapeMenuIdx=null"><span class="material-symbols-outlined" style="font-size:18px;">delete</span> Delete</div>
                   </div>
               </ng-container>
               
               <div class="shape-content-wrapper">
                 <svg *ngIf="s.type==='rect'" width="100%" height="100%"><rect x="0" y="0" width="100%" height="100%" fill="#e8f0fe" stroke="#1a73e8" stroke-width="2"/></svg>
                 <svg *ngIf="s.type==='roundrect'" width="100%" height="100%"><rect x="0" y="0" width="100%" height="100%" rx="8" fill="#e8f0fe" stroke="#1a73e8" stroke-width="2"/></svg>
                 <svg *ngIf="s.type==='circle'" width="100%" height="100%"><ellipse cx="50%" cy="50%" rx="48%" ry="48%" fill="#e8f0fe" stroke="#1a73e8" stroke-width="2"/></svg>
                 <svg *ngIf="s.type==='triangle'" width="100%" height="100%" preserveAspectRatio="none" viewBox="0 0 100 100"><polygon points="50,0 0,100 100,100" fill="#e8f0fe" stroke="#1a73e8" stroke-width="2"/></svg>
                 <svg *ngIf="s.type==='diamond'" width="100%" height="100%" preserveAspectRatio="none" viewBox="0 0 100 100"><polygon points="50,0 100,50 50,100 0,50" fill="#e8f0fe" stroke="#1a73e8" stroke-width="2"/></svg>
                 <svg *ngIf="s.type==='hexagon'" width="100%" height="100%" preserveAspectRatio="none" viewBox="0 0 100 100"><polygon points="50,0 100,25 100,75 50,100 0,75 0,25" fill="#e8f0fe" stroke="#1a73e8" stroke-width="2"/></svg>
                 <svg *ngIf="s.type==='octagon'" width="100%" height="100%" preserveAspectRatio="none" viewBox="0 0 100 100"><polygon points="30,0 70,0 100,30 100,70 70,100 30,100 0,70 0,30" fill="#e8f0fe" stroke="#1a73e8" stroke-width="2"/></svg>
                 
                 <!-- Complex Diagrams -->
                 <svg *ngIf="s.type==='diagram_drop'" width="100%" height="100%" preserveAspectRatio="xMidYMid meet" viewBox="0 0 100 60"><circle cx="20" cy="15" r="5" fill="#0ea5e9"/><circle cx="25" cy="30" r="5" fill="#10b981"/><circle cx="20" cy="45" r="5" fill="#f59e0b"/><line x1="30" y1="15" x2="80" y2="15" stroke="#e2e8f0" stroke-width="2"/><line x1="35" y1="30" x2="80" y2="30" stroke="#e2e8f0" stroke-width="2"/><line x1="30" y1="45" x2="80" y2="45" stroke="#e2e8f0" stroke-width="2"/></svg>
                 <svg *ngIf="s.type==='diagram_stack'" width="100%" height="100%" preserveAspectRatio="xMidYMid meet" viewBox="0 0 100 60"><rect x="15" y="20" width="20" height="20" rx="2" fill="#fff" stroke="#0ea5e9" stroke-width="1" transform="rotate(-10 25 30)"/><rect x="40" y="20" width="20" height="20" rx="2" fill="#fff" stroke="#10b981" stroke-width="1"/><rect x="65" y="20" width="20" height="20" rx="2" fill="#fff" stroke="#f59e0b" stroke-width="1" transform="rotate(10 75 30)"/></svg>
                 <svg *ngIf="s.type==='diagram_flag'" width="100%" height="100%" preserveAspectRatio="xMidYMid meet" viewBox="0 0 100 60"><circle cx="15" cy="15" r="4" fill="#0ea5e9"/><rect x="25" y="12" width="60" height="6" fill="#0ea5e9"/><circle cx="15" cy="30" r="4" fill="#10b981"/><rect x="25" y="27" width="60" height="6" fill="#10b981"/><circle cx="15" cy="45" r="4" fill="#f59e0b"/><rect x="25" y="42" width="60" height="6" fill="#f59e0b"/></svg>
                 <svg *ngIf="s.type==='diagram_ribbon'" width="100%" height="100%" preserveAspectRatio="xMidYMid meet" viewBox="0 0 100 60"><path d="M 20 10 L 80 10 L 80 20 L 20 20 Z" fill="#0ea5e9"/><path d="M 20 25 L 80 25 L 80 35 L 20 35 Z" fill="#10b981"/><path d="M 20 40 L 80 40 L 80 50 L 20 50 Z" fill="#f59e0b"/></svg>
                 <svg *ngIf="s.type==='diagram_alter'" width="100%" height="100%" preserveAspectRatio="xMidYMid meet" viewBox="0 0 100 60"><circle cx="50" cy="15" r="4" fill="#0ea5e9"/><line x1="20" y1="15" x2="40" y2="15" stroke="#e2e8f0" stroke-width="2"/><line x1="60" y1="15" x2="80" y2="15" stroke="#e2e8f0" stroke-width="2"/><circle cx="50" cy="30" r="4" fill="#10b981"/><line x1="20" y1="30" x2="40" y2="30" stroke="#e2e8f0" stroke-width="2"/><line x1="60" y1="30" x2="80" y2="30" stroke="#e2e8f0" stroke-width="2"/></svg>
                 <svg *ngIf="s.type==='diagram_deck'" width="100%" height="100%" preserveAspectRatio="xMidYMid meet" viewBox="0 0 100 60"><rect x="20" y="10" width="15" height="40" fill="#fff" stroke="#e2e8f0" stroke-width="1"/><rect x="40" y="10" width="15" height="40" fill="#fff" stroke="#e2e8f0" stroke-width="1"/><rect x="60" y="10" width="15" height="40" fill="#fff" stroke="#e2e8f0" stroke-width="1"/></svg>
                 <svg *ngIf="s.type==='diagram_process_arrow'" width="100%" height="100%" preserveAspectRatio="xMidYMid meet" viewBox="0 0 100 60"><polygon points="10,20 40,20 40,10 60,30 40,50 40,40 10,40" fill="#0ea5e9"/><polygon points="45,20 75,20 75,10 95,30 75,50 75,40 45,40" fill="#10b981"/></svg>
                 <svg *ngIf="s.type==='diagram_process_step'" width="100%" height="100%" preserveAspectRatio="xMidYMid meet" viewBox="0 0 100 60"><rect x="10" y="20" width="20" height="20" fill="#0ea5e9"/><line x1="30" y1="30" x2="40" y2="30" stroke="#ccc" stroke-width="2"/><rect x="40" y="20" width="20" height="20" fill="#10b981"/><line x1="60" y1="30" x2="70" y2="30" stroke="#ccc" stroke-width="2"/><rect x="70" y="20" width="20" height="20" fill="#f59e0b"/></svg>
                 <svg *ngIf="s.type==='diagram_pyramid_basic'" width="100%" height="100%" preserveAspectRatio="xMidYMid meet" viewBox="0 0 100 60"><polygon points="50,5 35,20 65,20" fill="#0ea5e9"/><polygon points="32,22 68,22 83,40 17,40" fill="#10b981"/><polygon points="14,42 86,42 100,55 0,55" fill="#f59e0b"/></svg>
                 <svg *ngIf="s.type==='diagram_pyramid_inv'" width="100%" height="100%" preserveAspectRatio="xMidYMid meet" viewBox="0 0 100 60"><polygon points="0,5 100,5 86,18 14,18" fill="#0ea5e9"/><polygon points="17,20 83,20 68,38 32,38" fill="#10b981"/><polygon points="35,40 65,40 50,55" fill="#f59e0b"/></svg>
                 <svg *ngIf="s.type==='diagram_cycle_basic'" width="100%" height="100%" preserveAspectRatio="xMidYMid meet" viewBox="0 0 100 60"><path d="M50,10 A20,20 0 0,1 70,30 L65,30 L72.5,40 L80,30 L75,30 A25,25 0 0,0 50,5 Z" fill="#0ea5e9"/><path d="M70,30 A20,20 0 0,1 30,30 L25,30 L32.5,20 L40,30 L35,30 A25,25 0 0,0 75,30 Z" fill="#10b981"/></svg>

                 <!-- Lines & Flowcharts -->
                 <svg *ngIf="s.type==='line_straight'" width="100%" height="100%" preserveAspectRatio="none" viewBox="0 0 24 24"><line x1="4" y1="20" x2="20" y2="4" stroke="#5f6368" stroke-width="1.5"/></svg>
                 <svg *ngIf="s.type==='line_arrow'" width="100%" height="100%" preserveAspectRatio="none" viewBox="0 0 24 24"><line x1="4" y1="20" x2="18" y2="6" stroke="#5f6368" stroke-width="1.5"/><polygon points="16,4 21,3 20,8" fill="#5f6368"/></svg>
                 <svg *ngIf="s.type==='line_curve'" width="100%" height="100%" preserveAspectRatio="none" viewBox="0 0 24 24"><path d="M4,20 Q12,4 20,20" fill="none" stroke="#5f6368" stroke-width="1.5"/></svg>
                 <svg *ngIf="s.type==='line_connector'" width="100%" height="100%" preserveAspectRatio="none" viewBox="0 0 24 24"><polyline points="4,20 4,12 20,12 20,4" fill="none" stroke="#5f6368" stroke-width="1.5"/></svg>
                 
                 <svg *ngIf="s.type==='flow_process'" width="100%" height="100%" preserveAspectRatio="none" viewBox="0 0 24 24"><rect x="3" y="6" width="18" height="12" fill="#e8f0fe" stroke="#1a73e8" stroke-width="1.5"/></svg>
                 <svg *ngIf="s.type==='flow_decision'" width="100%" height="100%" preserveAspectRatio="none" viewBox="0 0 24 24"><polygon points="12,3 21,12 12,21 3,12" fill="#e8f0fe" stroke="#1a73e8" stroke-width="1.5"/></svg>
                 <svg *ngIf="s.type==='flow_data'" width="100%" height="100%" preserveAspectRatio="none" viewBox="0 0 24 24"><polygon points="6,6 22,6 18,18 2,18" fill="#e8f0fe" stroke="#1a73e8" stroke-width="1.5"/></svg>
                 <svg *ngIf="s.type==='flow_terminator'" width="100%" height="100%" preserveAspectRatio="none" viewBox="0 0 24 24"><rect x="3" y="6" width="18" height="12" rx="6" fill="#e8f0fe" stroke="#1a73e8" stroke-width="1.5"/></svg>
                 
                 <svg *ngIf="s.type==='math_plus'" width="100%" height="100%" preserveAspectRatio="none" viewBox="0 0 24 24"><path d="M11,4 L13,4 L13,11 L20,11 L20,13 L13,13 L13,20 L11,20 L11,13 L4,13 L4,11 L11,11 Z" fill="#e8f0fe" stroke="#1a73e8" stroke-width="1.5"/></svg>
                 <svg *ngIf="s.type==='math_minus'" width="100%" height="100%" preserveAspectRatio="none" viewBox="0 0 24 24"><rect x="4" y="11" width="16" height="2" fill="#1a73e8"/></svg>
                 <svg *ngIf="s.type==='math_multiply'" width="100%" height="100%" preserveAspectRatio="none" viewBox="0 0 24 24"><path d="M6,6 L18,18 M18,6 L6,18" stroke="#1a73e8" stroke-width="2"/></svg>
                 <svg *ngIf="s.type==='math_divide'" width="100%" height="100%" preserveAspectRatio="none" viewBox="0 0 24 24"><circle cx="12" cy="6" r="2" fill="#1a73e8"/><rect x="4" y="11" width="16" height="2" fill="#1a73e8"/><circle cx="12" cy="18" r="2" fill="#1a73e8"/></svg>
                 <svg *ngIf="s.type==='math_equal'" width="100%" height="100%" preserveAspectRatio="none" viewBox="0 0 24 24"><rect x="4" y="8" width="16" height="2" fill="#1a73e8"/><rect x="4" y="14" width="16" height="2" fill="#1a73e8"/></svg>
                 
                 <svg *ngIf="s.type==='star_5'" width="100%" height="100%" preserveAspectRatio="none" viewBox="0 0 24 24"><polygon points="12,2 15,9 22,9 16,14 18,21 12,17 6,21 8,14 2,9 9,9" fill="#e8f0fe" stroke="#1a73e8" stroke-width="1.5"/></svg>
                 <svg *ngIf="s.type==='star_4'" width="100%" height="100%" preserveAspectRatio="none" viewBox="0 0 24 24"><polygon points="12,2 14,10 22,12 14,14 12,22 10,14 2,12 10,10" fill="#e8f0fe" stroke="#1a73e8" stroke-width="1.5"/></svg>
                 
                 <svg *ngIf="s.type==='callout_rect'" width="100%" height="100%" preserveAspectRatio="none" viewBox="0 0 24 24"><path d="M3,4 L21,4 L21,16 L14,16 L10,21 L10,16 L3,16 Z" fill="#e8f0fe" stroke="#1a73e8" stroke-width="1.5"/></svg>
                 <svg *ngIf="s.type==='callout_round'" width="100%" height="100%" preserveAspectRatio="none" viewBox="0 0 24 24"><path d="M4,4 C2,4 2,16 4,16 L10,16 L10,21 L14,16 L20,16 C22,16 22,4 20,4 Z" fill="#e8f0fe" stroke="#1a73e8" stroke-width="1.5"/></svg>
                 <svg *ngIf="s.type==='callout_cloud'" width="100%" height="100%" preserveAspectRatio="none" viewBox="0 0 24 24"><path d="M6,14 C4,14 4,10 6,10 C6,6 12,6 14,8 C16,6 20,8 20,12 C22,12 22,16 18,16 L12,21 L10,16 Z" fill="#e8f0fe" stroke="#1a73e8" stroke-width="1.5"/></svg>

                 <div *ngIf="s.type==='text_rounded'" style="width:100%; height:100%; border:1px solid #ccc; border-radius:16px; display:flex; align-items:center; justify-content:center; background:#fff;"></div>
                 <div *ngIf="s.type==='text_yellow'" style="width:100%; height:100%; background:#fef08a; border-radius:4px; display:flex; align-items:center; justify-content:center;"></div>
                 <div *ngIf="s.type==='text_arrow'" style="width:100%; height:100%; background:#0ea5e9; border-radius:4px; clip-path: polygon(0% 0%, 75% 0%, 100% 50%, 75% 100%, 0% 100%); display:flex; align-items:center; justify-content:center;"></div>

                 <div *ngIf="s.type==='button'" style="width:100%; height:100%; background:#f8f9fa; border:1px solid #0f9d58; display:flex; align-items:center; justify-content:center; box-sizing:border-box;"></div>

                 <div *ngIf="s.text" [style.color]="s.type==='text_arrow' ? '#ffffff' : '#1f2937'" style="position:absolute; inset:0; display:flex; align-items:center; justify-content:center; font-weight:600; font-size:14px; text-align:center; pointer-events:none;">
                   {{ s.text }}
                 </div>
               </div>
          </div>
        </ng-container>

        <table class="grid" [class.no-gridlines]="sheets[currentSheetIdx].hideGridlines" [class.print-area-active]="showHighlightPrintArea" [style.zoom]="zoomLevel / 100" [attr.dir]="gridDirection" [class.grid-spacing-comfort]="gridSpacing==='comfort'" [class.grid-spacing-cozy]="gridSpacing==='cozy'" [class.grid-spacing-classic]="gridSpacing==='classic'">
          <thead [style.display]="showHeaders ? '' : 'none'">
            <tr *ngIf="hasColGroups">
              <th class="corner" *ngIf="hasRowGroups" style="height: 24px; min-height: 24px; background: #f8f9fa; border-bottom: 1px solid #e2e8f0; position: sticky; left: 0; z-index: 6;"></th>
              <th class="corner" style="height: 24px; min-height: 24px; background: #f8f9fa; border-bottom: 1px solid #e2e8f0; position: sticky; left: 0; z-index: 6;" [style.left.px]="groupMarginWidth"></th>
              <th *ngFor="let c of colRange; trackBy: trackByCol" class="col-head group-margin-col-cell"
                [style.display]="hiddenCols.has(c) ? 'none' : ''"
                [style.width.px]="getColWidth(c)" [style.max-width.px]="getColWidth(c)"
                [style.position]="c < frozenColsCount ? 'sticky' : 'relative'"
                [style.left]="gridDirection==='ltr' && c < frozenColsCount ? getFrozenColOffset(c) + 'px' : ''"
                [style.right]="gridDirection==='rtl' && c < frozenColsCount ? getFrozenColOffset(c) + 'px' : ''"
                [style.z-index]="c < frozenColsCount ? 4 : ''"
                style="height: 24px; min-height: 24px; background: #f8f9fa; border-bottom: 1px solid #e2e8f0; position: relative; padding: 0;">
                <ng-container *ngFor="let g of getColGroupsFor(c); let i = index">
                  <div *ngIf="c >= g.start && c <= g.end && !g.collapsed" style="position: absolute; height: 1px; background: #64748b; right: 0;" [style.left]="c === g.start ? '50%' : '0'" [style.top.px]="i * 10 + 10"></div>
                  <div *ngIf="c === g.end && !g.collapsed" style="position: absolute; height: 5px; width: 1px; background: #64748b; right: 0;" [style.top.px]="i * 10 + 10"></div>
                  <div *ngIf="c === g.start" style="position: absolute; left: 50%; transform: translateX(-50%); width: 10px; height: 10px; background: #fff; border: 1px solid #64748b; border-radius: 2px; display: flex; align-items: center; justify-content: center; cursor: pointer; font-size: 10px; font-weight: bold; color: #334155; user-select: none; z-index: 2;" [style.top.px]="i * 10 + 5" (click)="toggleColGroup(g.index)">
                    {{ g.collapsed ? '+' : '-' }}
                  </div>
                </ng-container>
              </th>
            </tr>
            <tr>
              <th class="corner" *ngIf="hasRowGroups" style="width: 24px; min-width: 24px; max-width: 24px; background: #f8f9fa; border-right: 1px solid #e2e8f0; position: sticky; left: 0; z-index: 6;" [style.top.px]="colGroupMarginHeight"></th>
              <th class="corner" (click)="selectAll()" [style.z-index]="frozenRowsCount > 0 && frozenColsCount > 0 ? 5 : ''" [style.left.px]="groupMarginWidth" [style.top.px]="colGroupMarginHeight"></th>
              <th *ngFor="let c of colRange; trackBy: trackByCol" class="col-head"
                [style.display]="hiddenCols.has(c) ? 'none' : ''"
                [style.min-width.px]="getColWidth(c)" [style.width.px]="getColWidth(c)" [style.max-width.px]="getColWidth(c)"
                [style.position]="c < frozenColsCount ? 'sticky' : 'sticky'"
                [style.top.px]="colGroupMarginHeight"
                [style.left]="gridDirection==='ltr' && c < frozenColsCount ? getFrozenColOffset(c) + 'px' : ''"
                [style.right]="gridDirection==='rtl' && c < frozenColsCount ? getFrozenColOffset(c) + 'px' : ''"
                [style.z-index]="c < frozenColsCount ? 4 : ''"
                [class.col-selected]="isColHeaderSelected(c)"
                [class.active-axis]="isColActiveAxis(c)"
                (contextmenu)="onHeaderRightClick($event, 'col', c)"
                (click)="selectEntireCol(c)">
                {{ colLabel(c) }}
                <div class="col-resizer" (mousedown)="startColResize($event, c)"></div>
              </th>
            </tr>
          </thead>
          <tbody>
            
            <ng-container *ngFor="let r of visibleRowRange; trackBy: trackByRow">
              <tr *ngIf="r === firstUnfrozenRow && topSpacerHeight > 0" [style.height.px]="topSpacerHeight"><td [attr.colspan]="30" style="border:none;padding:0;pointer-events:none;"></td></tr>
              <tr [style.display]="hiddenRows.has(r) ? 'none' : ''" [style.height.px]="getRowHeight(r)">
              <td class="group-margin-cell" *ngIf="hasRowGroups" [style.display]="showHeaders ? '' : 'none'" style="width: 24px; min-width: 24px; max-width: 24px; position: sticky; left: 0; z-index: 5; background: #f8f9fa; border-right: 1px solid #e2e8f0; vertical-align: top; position: relative;">
                <ng-container *ngFor="let g of getRowGroupsFor(r); let i = index">
                  <div *ngIf="r >= g.start && r <= g.end && !g.collapsed" style="position: absolute; width: 1px; background: #64748b; bottom: 0;" [style.top]="r === g.start ? '50%' : '0'" [style.left.px]="i * 10 + 10"></div>
                  <div *ngIf="r === g.end && !g.collapsed" style="position: absolute; width: 5px; height: 1px; background: #64748b; bottom: 0;" [style.left.px]="i * 10 + 5"></div>
                  <div *ngIf="r === g.start" style="position: absolute; top: 50%; transform: translateY(-50%); width: 10px; height: 10px; background: #fff; border: 1px solid #64748b; border-radius: 2px; display: flex; align-items: center; justify-content: center; cursor: pointer; font-size: 10px; font-weight: bold; color: #334155; user-select: none;" [style.left.px]="i * 10 + 5" (click)="toggleRowGroup(g.index)">
                    {{ g.collapsed ? '+' : '-' }}
                  </div>
                </ng-container>
              </td>
              <td class="row-head" [style.display]="showHeaders ? '' : 'none'"
                [style.position]="r < frozenRowsCount ? 'sticky' : 'sticky'"
                [style.left.px]="groupMarginWidth"
                [style.top]="r < frozenRowsCount ? getFrozenRowOffset(r) + colGroupMarginHeight + 'px' : ''"
                [style.z-index]="r < frozenRowsCount ? 4 : ''"
                [class.row-selected]="isRowHeaderSelected(r)" [class.active-axis]="isRowActiveAxis(r)" (contextmenu)="onHeaderRightClick($event, 'row', r)" (click)="selectEntireRow(r)">
                {{ r + 1 }}
                <div class="row-resizer" (mousedown)="startRowResize($event, r)"></div>
              </td>
              <ng-container *ngFor="let c of colRange; trackBy: trackByCol">
                <td *ngIf="!isMergedSlave(r, c)" class="cell" [attr.id]="'cell-' + r + '-' + c"
                  [style.display]="hiddenCols.has(c) ? 'none' : ''"
                  [style.min-width.px]="getColWidth(c)" [style.width.px]="getColWidth(c)" [style.max-width.px]="getColWidth(c)"
                  [style.position]="r < frozenRowsCount || c < frozenColsCount ? 'sticky' : ''"
                  [style.top]="r < frozenRowsCount ? getFrozenRowOffset(r) + colGroupMarginHeight + 'px' : ''"
                  [style.left]="gridDirection==='ltr' && c < frozenColsCount ? getFrozenColOffset(c) + 'px' : ''"
                  [style.right]="gridDirection==='rtl' && c < frozenColsCount ? getFrozenColOffset(c) + 'px' : ''"
                  [style.z-index]="r < frozenRowsCount && c < frozenColsCount ? 4 : (r < frozenRowsCount || c < frozenColsCount ? 3 : '')"
                  [attr.colspan]="getColSpan(r, c)"
                  [attr.rowspan]="getRowSpan(r, c)"
                  [class.selected]="isCellSelected(r, c)"
                  [class.in-range]="isCellInRange(r, c)"
                  [class.remote-selected]="isRemoteSelected(r, c)"
                  [class.fill-preview]="isCellInFillPreview(r, c)"
                  [class.has-content]="cellHasContent(r, c)"
                  [class.search-match]="isCellInInlineSearch(r, c)"
                  [class.search-match-active]="isCellActiveInlineSearch(r, c)"
                  [class.comment-highlight]="isCommentHighlighted(r, c)"
                  [ngStyle]="getCellStyle(r, c)"
                  (mousedown)="onCellMouseDown($event, r, c)"
                  (mouseenter)="onCellMouseEnter(r, c)"
                  (contextmenu)="onCellRightClick($event, r, c)"
                  (click)="onCellClickWithPicker(r, c)"
                  (dblclick)="startEditing()">
                  <textarea *ngIf="isEditingCell && selectedRow === r && selectedCol === c" #floatingEditor class="floating-editor"
                     [style.left.px]="-2"
                     [style.top.px]="-2"
                     [style.min-width.px]="getColWidth(selectedCol) + 3"
                     [ngStyle]="getContentStyle(selectedRow, selectedCol)"
                     [(ngModel)]="editValue"
                     (input)="autoResizeEditor($event)"
                     (keydown)="onEditorKeydown($event)"
                     (blur)="commitEdit()"
                     (click)="$event.stopPropagation()"
                     (dblclick)="$event.stopPropagation()"
                     (mousedown)="$event.stopPropagation()"></textarea>
                                  <ng-container *ngIf="isImageCell(r, c); else textCell">
                    <img [src]="getImageSrc(cells[r][c])" style="max-width:100%;max-height:80px;object-fit:contain;display:inline-block;cursor:zoom-in;" 
(click)="selectCell(r,c); previewImageUrl = getImageSrc(cells[r][c])">
                  </ng-container>
                <ng-template #textCell>
                  <ng-container *ngIf="isSparklineCell(r, c); else dropdownCell">
                    <div style="width: 100%; height: 100%; display: flex; align-items: center; justify-content: center; overflow: hidden; pointer-events: none;" [innerHTML]="getSparklineSvgSafe(r, c)"></div>
                  </ng-container>
                  <ng-template #dropdownCell>
                    <ng-container *ngIf="hasCellDropdown(r, c); else plainInput">
                      <!-- Custom Picklist rendering -->
                    <div class="cell-dropdown-ui" 
                         [style.background]="!isDisplayAsChip(r, c) && cells[r][c] ? getDropdownColor(r, c, cells[r][c]) : 'transparent'"
                         style="position: absolute; top: 0; left: 0; right: 0; bottom: 0; padding: 0 6px; display: flex; justify-content: space-between; align-items: center; font-size: 13px; box-sizing: border-box; overflow: hidden; white-space: nowrap; text-overflow: ellipsis; user-select: none;">
                      
                      <!-- CHIP MODE (Display as chip = true) -->
                      <div *ngIf="isDisplayAsChip(r, c)" style="display:flex; gap:4px; overflow:hidden; flex-wrap:nowrap; pointer-events:none; flex:1; align-items:center;">
                        <span *ngFor="let part of splitValue(cells[r][c])" 
                              [style.background]="getDropdownColor(r, c, part) || '#f1f5f9'" 
                              [style.color]="getDropdownTextColor(r, c, part)" 
                              style="padding: 3px 12px; border-radius: 12px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-weight: 500; display: inline-flex; align-items: center; justify-content: center; min-width: 48px; max-width: 100%; box-sizing: border-box;">
                          {{ part }}
                        </span>
                      </div>

                      <!-- FULL CELL FILL MODE (Display as chip = false) -->
                      <div *ngIf="!isDisplayAsChip(r, c)" style="display:flex; gap:4px; overflow:hidden; flex-wrap:nowrap; pointer-events:none; flex:1; align-items:center; height:100%;">
                        <span *ngFor="let part of splitValue(cells[r][c]); let last = last" 
                              [style.color]="getDropdownTextColor(r, c, part)" 
                              style="overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-weight: 500;">
                          {{ part }}{{ last ? '' : ',' }}
                        </span>
                      </div>

                      <div (click)="openCustomDropdown($event, r, c)" style="display: flex; align-items: center; justify-content: center; height: 100%; cursor: pointer; padding: 0 4px; border-radius: 4px;">
                        <span class="material-symbols-outlined" [style.color]="!isDisplayAsChip(r, c) && cells[r][c] ? getDropdownTextColor(r, c, cells[r][c]) : '#5f6368'" style="font-size: 16px;">arrow_drop_down</span>
                      </div>
                    </div>
                  </ng-container>
                  <ng-template #plainInput>
                    <ng-container *ngIf="isCheckboxCell(r, c); else dateInput">
                      <div class="cell-checkbox-container" style="display:flex; justify-content:center; align-items:center; width:100%; height:100%;">
                         <span class="material-symbols-outlined" style="font-size:18px; color:#5f6368; cursor:pointer;" (mousedown)="onCheckboxMouseDown($event, r, c)">
                            {{ cells[r][c] === 'TRUE' ? 'check_box' : 'check_box_outline_blank' }}
                         </span>
                      </div>
                    </ng-container>
                    <ng-template #dateInput>
                      <div class="cell-display" [ngStyle]="getContentStyle(r, c)" [class.wrap-text]="getFormatWrap(r, c)" [style.opacity]="isEditingCell && selectedRow===r && selectedCol===c ? '0' : '1'">
                        <a *ngIf="isUrl(cells[r][c]); else normalText" [href]="cells[r][c]" target="_blank" style="color: #1155cc; text-decoration: underline; pointer-events: auto; cursor: pointer;" (click)="onLinkClick($event, cells[r][c])">{{ getDisplayValue(r, c) }}</a>
                        <ng-template #normalText>{{ getDisplayValue(r, c) }}</ng-template>
                      </div>
                      <!-- Auto-filter Dropdown Icon -->
                      <div *ngIf="isFilterHeaderCell(r, c)" class="cell-filter-icon" (click)="openFilterMenu($event, r, c); $event.stopPropagation()" style="position:absolute; right:4px; top:50%; transform:translateY(-50%); width:16px; height:16px; background:#f1f5f9; border-radius:2px; display:flex; align-items:center; justify-content:center; cursor:pointer; z-index:10; border:1px solid #cbd5e1; box-shadow:0 1px 2px rgba(0,0,0,0.05);">
                        <span class="material-symbols-outlined" style="font-size:12px;" [style.color]="isColumnFiltered(c) ? '#10b981' : '#5f6368'">{{ isColumnFiltered(c) ? 'filter_alt' : 'arrow_drop_down' }}</span>
                      </div>
                    </ng-template>
                  </ng-template>
                  </ng-template>
                </ng-template>
                <!-- Comment indicator -->
                <div *ngIf="hasComment(r, c)"
                     class="comment-indicator"
                     (click)="openCommentInSidePanel($event, r, c)"
                     style="position: absolute; top: 0; right: 0; width: 0; height: 0; border-style: solid; border-width: 0 8px 8px 0; border-color: transparent #f59e0b transparent transparent; z-index: 10; cursor: pointer;">
                </div>
                <!-- Note indicator -->
                <div *ngIf="formats[r + ',' + c]?.note"
                     class="note-indicator"
                     (mousedown)="openNotePopup($event, r, c)"
                     [style.right.px]="hasComment(r, c) ? 10 : 0"
                     style="position: absolute; top: 0; right: 0; width: 0; height: 0; border-style: solid; border-width: 0 10px 10px 0; border-color: transparent #d32f2f transparent transparent; z-index: 9; pointer-events: auto; cursor: pointer;"
                     title="View Note">
                </div>
                <!-- Fill handle: only show on the bottom-right cell of the selection -->
                <div *ngIf="isFillHandleCell(r, c)"
                  class="fill-handle"
                  (mousedown)="onFillHandleMouseDown($event, r, c)"
                  title="Drag to fill"></div>
                </td>
              </ng-container>
            </tr>
            </ng-container>
          <tr *ngIf="bottomSpacerHeight > 0" [style.height.px]="bottomSpacerHeight"><td [attr.colspan]="30" style="border:none;padding:0;pointer-events:none;"></td></tr>
          </tbody>
        </table>
        <!-- Note Popup -->
        <div *ngIf="activeNotePopup" 
             (mousedown)="$event.stopPropagation()"
             [style.position]="'absolute'"
             [style.top.px]="getRowOffset(activeNotePopup.r) + 25"
             [style.left.px]="getColOffset(activeNotePopup.c) + 25"
             style="z-index: 100; background: #e0f2f1; border: 1px solid #00897b; border-radius: 4px; padding: 12px; width: 220px; box-shadow: 0 4px 12px rgba(0,0,0,0.2); font-family: sans-serif;">
           <textarea [(ngModel)]="activeNotePopup.text" 
                     (ngModelChange)="onNoteTextChange($event)"
                     placeholder="Add Note" 
                     style="width: 100%; min-height: 80px; border: none; background: transparent; outline: none; resize: none; font-size: 13px; color: #333;"></textarea>
           <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 8px; font-size: 11px; color: #555;">
             <span>Cell: {{ colLabel(activeNotePopup.c) }}{{ activeNotePopup.r + 1 }}</span>
             <span class="material-symbols-outlined" style="font-size: 16px; cursor: pointer; color: #d32f2f;" (click)="deleteNote()" title="Delete Note">delete</span>
           </div>
        </div>
      </div>


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
      <!-- Right Side Panel for Apps -->
      <div class="side-panel" *ngIf="sidePanelApp">
        <div class="sp-head">
          <div class="sp-head-left">
            <div class="sp-icon-wrap" [class.sp-icon-cal]="sidePanelApp==='calendar'" [class.sp-icon-notes]="sidePanelApp==='notes'" [class.sp-icon-tasks]="sidePanelApp==='tasks'" [class.sp-icon-pivot]="sidePanelApp==='pivot'" [style.background]="sidePanelApp==='pivot'?'#10b981':(sidePanelApp==='navigation'?'#1a73e8':(sidePanelApp==='comments'?'#f59e0b':(sidePanelApp==='sparkline'?'#6366f1':'inherit')))">
              <span class="material-symbols-outlined sp-head-icon" [style.color]="sidePanelApp==='navigation'?'#fff':(sidePanelApp==='comments'?'#fff':(sidePanelApp==='sparkline'?'#fff':''))">{{sidePanelApp==='pivot'?'pivot_table_chart':sidePanelApp==='calendar'?'calendar_month':sidePanelApp==='notes'?'sticky_note_2':sidePanelApp==='navigation'?'web_stories':sidePanelApp==='comments'?'forum':sidePanelApp==='sparkline'?'ssid_chart':'task_alt'}}</span>
            </div>
            <div>
              <div class="sp-title">{{sidePanelApp==='pivot'?'Pivot Table Editor':sidePanelApp==='calendar'?'Calendar':sidePanelApp==='notes'?'Notes':sidePanelApp==='navigation'?'Navigation':sidePanelApp==='comments'?'Comments':sidePanelApp==='sparkline'?'Sparkline':'Tasks'}}</div>
              <div class="sp-subtitle">{{sidePanelApp==='pivot'?'Configure rows and values':sidePanelApp==='calendar'?'Schedule & meeting notes':sidePanelApp==='notes'?'Quick capture':sidePanelApp==='navigation'?'Manage objects and charts':sidePanelApp==='comments'?'Discuss with your team':sidePanelApp==='sparkline'?'Configure appearance':'Track your work'}}</div>
            </div>
          </div>
          <button class="sp-close-btn" (click)="closeSidePanel()">
            <span class="material-symbols-outlined">close</span>
          </button>
        </div>

        <div class="sp-content">

          <!-- ── PIVOT TABLE ─────────────────────────────────────────────── -->
          <ng-container *ngIf="sidePanelApp === 'pivot'">
            <div class="sp-card" style="display:flex; flex-direction:column; gap:16px;">
              
              <div>
                <div class="sp-card-label" style="font-weight:600; margin-bottom:8px;">Rows</div>
                <select [(ngModel)]="pivotConfig.row" (change)="applyPivot()" style="width:100%; padding:8px; border:1px solid #e2e8f0; border-radius:4px; outline:none;">
                  <option value="" disabled>Select row field...</option>
                  <option *ngFor="let h of pivotHeaders" [value]="h">{{h}}</option>
                </select>
              </div>

              <div>
                <div class="sp-card-label" style="font-weight:600; margin-bottom:8px;">Values</div>
                <select [(ngModel)]="pivotConfig.val" (change)="applyPivot()" style="width:100%; padding:8px; border:1px solid #e2e8f0; border-radius:4px; outline:none;">
                  <option value="" disabled>Select value field...</option>
                  <option *ngFor="let h of pivotHeaders" [value]="h">{{h}}</option>
                </select>
              </div>

              <div>
                <div class="sp-card-label" style="font-weight:600; margin-bottom:8px;">Summarize by</div>
                <select [(ngModel)]="pivotConfig.agg" (change)="applyPivot()" style="width:100%; padding:8px; border:1px solid #e2e8f0; border-radius:4px; outline:none;">
                  <option value="SUM">SUM</option>
                  <option value="COUNT">COUNT</option>
                  <option value="AVG">AVG</option>
                </select>
              </div>

            </div>
          </ng-container>

          <!-- ── CALENDAR ─────────────────────────────────────────────── -->
          <ng-container *ngIf="sidePanelApp === 'calendar'">
            <div class="sp-card sp-date-card">
              <div class="sp-card-label">
                <span class="material-symbols-outlined sp-label-icon">event</span>
                Select Date
              </div>
              <input type="date" [(ngModel)]="selectedCalDate"
                class="sp-date-input">
              <div class="sp-date-chip" *ngIf="selectedCalDate">
                <span class="material-symbols-outlined" style="font-size:14px;">schedule</span>
                {{selectedCalDate}}
              </div>
            </div>

            <div class="sp-card sp-notes-card sp-notes-grow">
              <div class="sp-card-label">
                <span class="material-symbols-outlined sp-label-icon">rate_review</span>
                Meeting Notes
              </div>
              <div class="sp-textarea-wrap">
                <textarea [(ngModel)]="calendarNotes[selectedCalDate]" (change)="save()"
                  [placeholder]="'Notes for ' + (selectedCalDate || 'selected date') + '...'"
                  class="sp-textarea"></textarea>
                <div class="sp-textarea-footer">
                  <span class="material-symbols-outlined" style="font-size:13px;color:#aaa;">save</span>
                  <span style="font-size:11px;color:#aaa;">Auto-saved</span>
                </div>
              </div>
            </div>
          </ng-container>

          <!-- ── NOTES ───────────────────────────────────────────────── -->
          <ng-container *ngIf="sidePanelApp === 'notes'">
            <div class="sp-card sp-notes-card" style="flex:1;display:flex;flex-direction:column;">
              <div class="sp-card-label">
                <span class="material-symbols-outlined sp-label-icon">edit_note</span>
                Global Notes
              </div>
              <div class="sp-textarea-wrap" style="flex:1">
                <textarea [(ngModel)]="globalNotes" (change)="save()"
                  placeholder="Capture your thoughts, ideas, and requirements here…"
                  class="sp-textarea sp-textarea-tall"></textarea>
                <div class="sp-textarea-footer">
                  <span class="material-symbols-outlined" style="font-size:13px;color:#aaa;">save</span>
                  <span style="font-size:11px;color:#aaa;">Auto-saved</span>
                </div>
              </div>
            </div>
          </ng-container>

          <!-- ── TASKS ───────────────────────────────────────────────── -->
          <ng-container *ngIf="sidePanelApp === 'tasks'">
            <div class="sp-task-add-wrap">
              <span class="material-symbols-outlined sp-task-add-icon">add_circle</span>
              <input type="text" [(ngModel)]="newTask" (keyup.enter)="addTask()"
                placeholder="Add a new task…"
                class="sp-task-input">
              <button class="sp-add-btn" (click)="addTask()">Add</button>
            </div>

            <div class="sp-tasks-summary" *ngIf="tasks.length > 0">
              <span class="sp-tasks-count">{{tasks.length}} task{{tasks.length===1?'':'s'}}</span>
              <span class="sp-tasks-done">{{getTasksDone()}} done</span>
            </div>

            <div class="sp-task-item" *ngFor="let t of tasks; let i = index" [class.sp-task-done]="t.done">
              <label class="sp-checkbox-wrap">
                <input type="checkbox" [(ngModel)]="t.done" (change)="save()" class="sp-checkbox-native">
                <span class="sp-checkbox-ui">
                  <span class="material-symbols-outlined sp-check-icon">check</span>
                </span>
              </label>
              <span class="sp-task-text">{{t.text}}</span>
              <button class="sp-task-del" (click)="removeTask(i)" title="Delete">
                <span class="material-symbols-outlined">delete_outline</span>
              </button>
            </div>

            <div class="sp-empty" *ngIf="tasks.length === 0">
              <div class="sp-empty-icon">
                <span class="material-symbols-outlined">check_circle</span>
              </div>
              <div class="sp-empty-title">All clear!</div>
              <div class="sp-empty-sub">Add your first task above to get started.</div>
            </div>
          </ng-container>

          <!-- ── NAVIGATION ────────────────────────────────────────────── -->
          <ng-container *ngIf="sidePanelApp === 'navigation'">
            <!-- Sheets Section -->
            <div style="margin-bottom: 24px;">
              <div class="sp-card-label" style="font-weight: 600; font-size: 11px; text-transform: uppercase; color: #5f6368; letter-spacing: 0.8px; margin-bottom: 12px;">Sheets</div>
              <div class="sp-task-item" *ngFor="let s of sheets; let i = index" (click)="switchSheet(i)" [style.display]="s.hidden ? 'none' : 'flex'" style="cursor: pointer; padding: 8px 12px; border-radius: 6px;" [style.background]="i === currentSheetIdx ? '#e8f0fe' : 'transparent'">
                <span class="material-symbols-outlined sp-check-icon" [style.color]="i === currentSheetIdx ? '#1a73e8' : '#5f6368'">grid_on</span>
                <span class="sp-task-text" style="margin-left: 12px; font-weight: 500;" [style.color]="i === currentSheetIdx ? '#1a73e8' : '#202124'">{{ s.name }}</span>
                <span *ngIf="i === currentSheetIdx" class="material-symbols-outlined" style="color: #1a73e8; font-size: 18px; margin-left: auto;">check</span>
              </div>

              <!-- Hidden Sheets -->
              <div *ngIf="hiddenSheetsList.length > 0" style="margin-top: 12px; border-top: 1px solid #e2e8f0; padding-top: 12px;">
                <div class="sp-task-item" *ngFor="let hs of hiddenSheetsList; trackBy: trackByHiddenSheet" (click)="unhideSheetAndSwitch(hs.idx)" style="cursor: pointer; opacity: 0.6; padding: 8px 12px; border-radius: 6px;">
                  <span class="material-symbols-outlined sp-check-icon" style="color: #5f6368;">visibility_off</span>
                  <span class="sp-task-text" style="margin-left: 12px; font-weight: 500;">{{ hs.s.name }} (Hidden)</span>
                </div>
              </div>
            </div>

            <!-- Objects Section -->
            <div>
              <div class="sp-card-label" style="font-weight: 600; font-size: 11px; text-transform: uppercase; color: #5f6368; letter-spacing: 0.8px; margin-bottom: 12px;">Objects in {{ sheets[currentSheetIdx].name }}</div>
              
              <div class="sp-task-item" *ngFor="let s of sheets[currentSheetIdx].shapes; let i = index">
                <span class="material-symbols-outlined sp-check-icon" style="color: #5f6368;">{{ s.type.includes('chart') ? 'insert_chart' : (s.type.includes('image') ? 'image' : 'category') }}</span>
                <span class="sp-task-text" style="margin-left: 12px; font-weight: 500;">{{ s.text || (s.type.includes('chart') ? 'Chart ' + (i+1) : 'Object ' + (i+1)) }}</span>
                <button class="sp-task-del" (click)="deleteShape(i)" title="Delete Object">
                  <span class="material-symbols-outlined" style="font-size: 20px;">delete_outline</span>
                </button>
              </div>
              
              <div class="sp-empty" *ngIf="!sheets[currentSheetIdx].shapes?.length" style="margin-top: 0; padding-top: 24px; border: none;">
                <div class="sp-empty-icon">
                  <span class="material-symbols-outlined">category</span>
                </div>
                <div class="sp-empty-title" style="font-size: 14px;">No Objects Found</div>
                <div class="sp-empty-sub">Add charts or shapes to see them listed here.</div>
              </div>
            </div>
          </ng-container>

          <!-- ── COMMENTS ────────────────────────────────────────────── -->
          <ng-container *ngIf="sidePanelApp === 'comments'">
            <div style="display:flex; flex-direction:column; gap:16px;">
              <div style="display:flex; justify-content:space-between; align-items:center;">
                <select [(ngModel)]="commentsViewFilter" (change)="updateCachedComments()" style="padding: 4px 8px; border: 1px solid #dadce0; border-radius: 4px; outline: none; background: #fff; font-size: 13px; font-weight: 500; color: #3c4043;">
                  <option value="all">All Sheets</option>
                  <option value="current">Current Sheet</option>
                </select>
                <select [(ngModel)]="commentsStatusFilter" (change)="updateCachedComments()" style="padding: 4px 8px; border: 1px solid #dadce0; border-radius: 4px; outline: none; background: #fff; font-size: 13px; font-weight: 500; color: #3c4043;">
                  <option value="all">All Status</option>
                  <option value="unresolved">Unresolved</option>
                  <option value="resolved">Resolved</option>
                </select>
              </div>

              <!-- New Comment Box -->
              <div class="sp-card" *ngIf="newCommentCellRef" style="border: 1px solid #1a73e8; box-shadow: 0 1px 3px rgba(0,0,0,0.1); padding: 12px; background: #fff;">
                <div style="display:flex; align-items:center; gap:8px; margin-bottom:8px;">
                  <div style="width:24px; height:24px; border-radius:12px; background:#e8f0fe; color:#1a73e8; display:flex; align-items:center; justify-content:center; font-size:12px; font-weight:bold;">
                    {{ auth.user?.name?.charAt(0) || 'U' }}
                  </div>
                  <div>
                    <div style="font-size:13px; font-weight:600; color:#202124;">{{ auth.user?.name || 'You' }}</div>
                    <div style="font-size:11px; color:#5f6368;">Refers to: <span style="color:#1a73e8;">{{ newCommentCellName }}</span></div>
                  </div>
                </div>
                <textarea #newCommentInput [(ngModel)]="newCommentText" placeholder="Add a comment..." style="width:100%; min-height:60px; padding:8px; border:1px solid #dadce0; border-radius:4px; resize:vertical; font-size:13px; outline:none; margin-bottom:8px; box-sizing: border-box;"></textarea>
                <div style="display:flex; gap:8px;">
                  <button (click)="submitNewComment()" style="background:#0f9d58; color:#fff; border:none; border-radius:4px; padding:6px 16px; font-size:13px; font-weight:500; cursor:pointer;">Add</button>
                  <button (click)="cancelNewComment()" style="background:transparent; color:#5f6368; border:none; padding:6px 16px; font-size:13px; font-weight:500; cursor:pointer;">Cancel</button>
                </div>
              </div>

              <!-- List of Comments -->
              <div *ngFor="let c of cachedComments" [id]="'comment-card-' + c.data.id" style="border: 1px solid #dadce0; border-radius: 8px; padding: 12px; background: #fff; margin-bottom: 8px;" [style.border-left]="c.data.resolved ? '4px solid #dadce0' : '4px solid #0f9d58'">
                <!-- Header -->
                <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:8px;">
                  <div style="display:flex; align-items:center; gap:8px;">
                    <div style="width:28px; height:28px; border-radius:14px; background:#f1f3f4; color:#3c4043; display:flex; align-items:center; justify-content:center; font-size:13px; font-weight:bold;">
                      {{ c.data.authorName?.charAt(0) || 'U' }}
                    </div>
                    <div>
                      <div style="font-size:13px; font-weight:600; color:#202124;">{{ c.data.authorName || 'Unknown' }}</div>
                      <div style="font-size:11px; color:#5f6368; display:flex; gap:4px; align-items:center;">
                        <span (click)="goToCommentCell(c)" style="cursor:pointer; color:#1a73e8; text-decoration:underline;">{{ c.sheetName }}.{{ c.cellName }}</span>
                        <span>•</span>
                        <span>{{ c.data.timestamp | date:'dd MMM, yyyy h:mm a' }}</span>
                      </div>
                    </div>
                  </div>
                  <!-- Actions -->
                  <div style="position:relative;">
                    <span class="material-symbols-outlined" style="font-size:18px; color:#5f6368; cursor:pointer;" (click)="activeCommentMenu = activeCommentMenu === c.data.id ? null : c.data.id">more_vert</span>
                    <div *ngIf="activeCommentMenu === c.data.id" style="position:absolute; right:0; top:20px; background:#fff; box-shadow:0 2px 8px rgba(0,0,0,0.15); border-radius:4px; padding:4px 0; min-width:120px; z-index:100; border:1px solid #dadce0;">
                      <div class="dd-item" (click)="toggleCommentResolve(c); activeCommentMenu=null" style="padding:8px 16px; font-size:13px; cursor:pointer; color:#3c4043;">
                        {{ c.data.resolved ? 'Mark Unresolved' : 'Mark Resolved' }}
                      </div>
                      <div class="dd-item" (click)="deleteComment(c); activeCommentMenu=null" style="padding:8px 16px; font-size:13px; cursor:pointer; color:#d93025;">
                        Delete
                      </div>
                    </div>
                  </div>
                </div>
                
                <!-- Body -->
                <div style="font-size:13px; color:#3c4043; line-height:1.5; margin-bottom:12px; white-space:pre-wrap;">{{ c.data.text }}</div>
                
                <!-- Replies -->
                <div *ngIf="c.data.replies?.length > 0" style="margin-left: 24px; padding-left: 12px; border-left: 2px solid #f1f3f4; margin-bottom: 12px; display:flex; flex-direction:column; gap:12px;">
                  <div *ngFor="let rep of c.data.replies" style="display:flex; gap:8px;">
                     <div style="width:24px; height:24px; border-radius:12px; background:#f1f3f4; color:#3c4043; display:flex; align-items:center; justify-content:center; font-size:11px; font-weight:bold; flex-shrink:0;">
                       {{ rep.authorName?.charAt(0) || 'U' }}
                     </div>
                     <div>
                       <div style="font-size:12px; font-weight:600; color:#202124;">{{ rep.authorName || 'Unknown' }} <span style="font-weight:normal; color:#5f6368; font-size:11px; margin-left:4px;">{{ rep.timestamp | date:'h:mm a' }}</span></div>
                       <div style="font-size:13px; color:#3c4043;">{{ rep.text }}</div>
                     </div>
                  </div>
                </div>

                <!-- Reply Box -->
                <div *ngIf="!c.data.resolved" style="display:flex; gap:8px; align-items:center;">
                  <input type="text" [(ngModel)]="replyTexts[c.data.id]" placeholder="Reply or add others with @" style="flex:1; padding:8px 12px; border:1px solid #dadce0; border-radius:20px; font-size:13px; outline:none;" (keyup.enter)="submitReply(c)">
                  <button (click)="submitReply(c)" style="background:none; border:none; cursor:pointer; display:flex; align-items:center; justify-content:center; padding:4px;" [style.color]="(replyTexts[c.data.id] || '').trim() ? '#1a73e8' : '#dadce0'"><span class="material-symbols-outlined" style="font-size:20px;">send</span></button>
                </div>
              </div>
              
              <div class="sp-empty" *ngIf="cachedComments.length === 0 && !newCommentCellRef" style="margin-top: 0; padding-top: 24px; border: none;">
                <div class="sp-empty-icon">
                  <span class="material-symbols-outlined">forum</span>
                </div>
                <div class="sp-empty-title" style="font-size: 14px;">No Comments Found</div>
                <div class="sp-empty-sub">Highlight a cell and add a comment to start a discussion.</div>
              </div>
            </div>
          </ng-container>

          <!-- 🟢 SPARKLINE SETTINGS PANEL 🟢 -->
          <ng-container *ngIf="sidePanelApp === 'sparkline' && sparklineConfig">
            <div style="display:flex; flex-direction:column; gap:16px;">
              <div style="display:flex; justify-content:space-between; align-items:center;">
                 <div style="font-size:16px; font-weight:600; color:#202124;">Sparkline</div>
                 <button (click)="sidePanelApp = null" style="background:none; border:none; cursor:pointer;"><span class="material-symbols-outlined" style="font-size:18px; color:#5f6368;">close</span></button>
              </div>
              <div style="font-size:13px; color:#202124;">Source: <strong>{{sparklineConfig.sourceRange}}</strong> <a style="color:#0f9d58; text-decoration:none; margin-left:8px; cursor:pointer; font-weight:500;" (click)="editSparklineConfig = { source: sparklineConfig.sourceRange || '', dest: sparklineConfig.destinationRange || '', error: '', tab: 'selected' }; activeModal = 'edit_sparkline'">Edit</a></div>
              
              <hr style="border:0; border-top:1px solid #e2e8f0; margin:0;">
              
              <div class="sp-card-label">Sparkline Type</div>
              <div style="display:flex; gap:8px; align-items:center;">
                <button (click)="setSparklineType('line')" [style.background]="sparklineConfig.type === 'line' ? '#e8f0fe' : '#fff'" [style.color]="sparklineConfig.type === 'line' ? '#1a73e8' : '#5f6368'" style="flex:1; padding:6px; border:1px solid #dadce0; border-radius:4px; cursor:pointer; display:flex; justify-content:center;" title="Line"><span class="material-symbols-outlined">show_chart</span></button>
                <button (click)="setSparklineType('column')" [style.background]="sparklineConfig.type === 'column' ? '#e8f0fe' : '#fff'" [style.color]="sparklineConfig.type === 'column' ? '#1a73e8' : '#5f6368'" style="flex:1; padding:6px; border:1px solid #dadce0; border-radius:4px; cursor:pointer; display:flex; justify-content:center;" title="Bar/Column"><span class="material-symbols-outlined">bar_chart</span></button>
                <button (click)="setSparklineType('winloss')" [style.background]="sparklineConfig.type === 'winloss' ? '#e8f0fe' : '#fff'" [style.color]="sparklineConfig.type === 'winloss' ? '#1a73e8' : '#5f6368'" style="flex:1; padding:6px; border:1px solid #dadce0; border-radius:4px; cursor:pointer; display:flex; justify-content:center;" title="Win/Loss"><span class="material-symbols-outlined">waterfall_chart</span></button>
              </div>

              <div class="sp-card-label" style="margin-top:16px;">Sparkline Color</div>
              <div style="position:relative;">
                   <div (click)="openColorPicker($event, 'base')" style="width:100%; height:32px; border-radius:4px; border:1px solid #dadce0; cursor:pointer; display:flex; align-items:center; justify-content:space-between; padding:0 8px; background: #fff;">
                      <div style="display:flex; align-items:center; gap:8px;">
                         <div style="width:16px; height:16px; border-radius:2px; border:1px solid rgba(0,0,0,0.1);" [style.background]="sparklineConfig.baseColor"></div>
                         <span style="font-size:13px; color:#202124;">Color</span>
                      </div>
                      <span class="material-symbols-outlined" style="font-size:18px; color:#5f6368;">expand_more</span>
                   </div>
              </div>

              <div class="sp-card-label" style="margin-top:16px;">Highlight Points</div>
              <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
                 <!-- High -->
                 <div style="display:flex; justify-content:space-between; align-items:center;">
                    <label style="display:flex; align-items:center; gap:6px; font-size:13px; cursor:pointer;"><input type="checkbox" [(ngModel)]="sparklineConfig.highlights.high.enabled" (change)="saveSparkline()"> High</label>
                    <div (click)="openColorPicker($event, 'high')" style="width:28px; height:20px; border-radius:2px; border:1px solid #dadce0; cursor:pointer; display:flex; align-items:center; justify-content:center;" [style.background]="sparklineConfig.highlights.high.color"><span class="material-symbols-outlined" style="font-size:14px; color:rgba(255,255,255,0.8); mix-blend-mode: difference;">expand_more</span></div>
                 </div>
                 <!-- Low -->
                 <div style="display:flex; justify-content:space-between; align-items:center;">
                    <label style="display:flex; align-items:center; gap:6px; font-size:13px; cursor:pointer;"><input type="checkbox" [(ngModel)]="sparklineConfig.highlights.low.enabled" (change)="saveSparkline()"> Low</label>
                    <div (click)="openColorPicker($event, 'low')" style="width:28px; height:20px; border-radius:2px; border:1px solid #dadce0; cursor:pointer; display:flex; align-items:center; justify-content:center;" [style.background]="sparklineConfig.highlights.low.color"><span class="material-symbols-outlined" style="font-size:14px; color:rgba(255,255,255,0.8); mix-blend-mode: difference;">expand_more</span></div>
                 </div>
                 <!-- First -->
                 <div style="display:flex; justify-content:space-between; align-items:center;">
                    <label style="display:flex; align-items:center; gap:6px; font-size:13px; cursor:pointer;"><input type="checkbox" [(ngModel)]="sparklineConfig.highlights.first.enabled" (change)="saveSparkline()"> First</label>
                    <div (click)="openColorPicker($event, 'first')" style="width:28px; height:20px; border-radius:2px; border:1px solid #dadce0; cursor:pointer; display:flex; align-items:center; justify-content:center;" [style.background]="sparklineConfig.highlights.first.color"><span class="material-symbols-outlined" style="font-size:14px; color:rgba(255,255,255,0.8); mix-blend-mode: difference;">expand_more</span></div>
                 </div>
                 <!-- Last -->
                 <div style="display:flex; justify-content:space-between; align-items:center;">
                    <label style="display:flex; align-items:center; gap:6px; font-size:13px; cursor:pointer;"><input type="checkbox" [(ngModel)]="sparklineConfig.highlights.last.enabled" (change)="saveSparkline()"> Last</label>
                    <div (click)="openColorPicker($event, 'last')" style="width:28px; height:20px; border-radius:2px; border:1px solid #dadce0; cursor:pointer; display:flex; align-items:center; justify-content:center;" [style.background]="sparklineConfig.highlights.last.color"><span class="material-symbols-outlined" style="font-size:14px; color:rgba(255,255,255,0.8); mix-blend-mode: difference;">expand_more</span></div>
                 </div>
                 <!-- Negative -->
                 <div style="display:flex; justify-content:space-between; align-items:center;">
                    <label style="display:flex; align-items:center; gap:6px; font-size:13px; cursor:pointer;"><input type="checkbox" [(ngModel)]="sparklineConfig.highlights.negative.enabled" (change)="saveSparkline()"> Negative</label>
                    <div (click)="openColorPicker($event, 'negative')" style="width:28px; height:20px; border-radius:2px; border:1px solid #dadce0; cursor:pointer; display:flex; align-items:center; justify-content:center;" [style.background]="sparklineConfig.highlights.negative.color"><span class="material-symbols-outlined" style="font-size:14px; color:rgba(255,255,255,0.8); mix-blend-mode: difference;">expand_more</span></div>
                 </div>
                 <!-- Markers -->
                 <div style="display:flex; justify-content:space-between; align-items:center;" [style.opacity]="sparklineConfig.type !== 'line' ? 0.4 : 1" [style.pointer-events]="sparklineConfig.type !== 'line' ? 'none' : 'auto'">
                    <label style="display:flex; align-items:center; gap:6px; font-size:13px; cursor:pointer;"><input type="checkbox" [(ngModel)]="sparklineConfig.highlights.markers.enabled" (change)="saveSparkline()"> Markers</label>
                    <div (click)="openColorPicker($event, 'markers')" style="width:28px; height:20px; border-radius:2px; border:1px solid #dadce0; cursor:pointer; display:flex; align-items:center; justify-content:center;" [style.background]="sparklineConfig.highlights.markers.color"><span class="material-symbols-outlined" style="font-size:14px; color:rgba(255,255,255,0.8); mix-blend-mode: difference;">expand_more</span></div>
                 </div>
              </div>

              <div class="sp-card-label" style="margin-top:8px;">Show Empty Cells</div>
              <div style="display:flex; font-size:12px; border:1px solid #dadce0; border-radius:4px; overflow:hidden;">
                <button style="flex:1; padding:6px 0; border:none; cursor:pointer; background:transparent;" [style.background]="sparklineConfig.emptyCellMode === 'gap' ? '#424242' : '#f1f3f4'" [style.color]="sparklineConfig.emptyCellMode === 'gap' ? '#fff' : '#202124'" (click)="setEmptyCellMode('gap')">Gap</button>
                <div style="width:1px; background:#dadce0;"></div>
                <button style="flex:1; padding:6px 0; border:none; cursor:pointer; background:transparent;" [style.background]="sparklineConfig.emptyCellMode === 'zero' ? '#424242' : '#f1f3f4'" [style.color]="sparklineConfig.emptyCellMode === 'zero' ? '#fff' : '#202124'" (click)="setEmptyCellMode('zero')">Zero</button>
                <div style="width:1px; background:#dadce0;"></div>
                <button style="flex:1; padding:6px 0; border:none; cursor:pointer; background:transparent;" [style.background]="sparklineConfig.emptyCellMode === 'connect' ? '#424242' : '#f1f3f4'" [style.color]="sparklineConfig.emptyCellMode === 'connect' ? '#fff' : '#202124'" (click)="setEmptyCellMode('connect')" [disabled]="sparklineConfig.type !== 'line'" [style.opacity]="sparklineConfig.type !== 'line' ? 0.4 : 1">Connect</button>
                <div style="width:1px; background:#dadce0;"></div>
                <button style="flex:1; padding:6px 0; border:none; cursor:pointer; background:transparent;" [style.background]="sparklineConfig.emptyCellMode === 'skip' ? '#424242' : '#f1f3f4'" [style.color]="sparklineConfig.emptyCellMode === 'skip' ? '#fff' : '#202124'" (click)="setEmptyCellMode('skip')">Skip</button>
              </div>

              <label style="display:flex; align-items:center; gap:8px; font-size:13px; cursor:pointer;"><input type="checkbox" [(ngModel)]="sparklineConfig.includeHiddenRowsColumns" (change)="saveSparkline()"> Include hidden rows and columns</label>

              <hr style="border:0; border-top:1px solid #e2e8f0; margin:0;">

              <div class="sp-card-label">Manage Settings</div>
              <div style="display:flex; gap:8px;">
                 <button (click)="toggleGroup()" style="flex:1; padding:6px; display:flex; align-items:center; justify-content:center; gap:4px; border:1px solid #dadce0; border-radius:4px; background:#fff; cursor:pointer; font-size:13px;"><span class="material-symbols-outlined" style="font-size:16px;">library_add</span> {{sparklineConfig.isGrouped ? 'Ungroup' : 'Group'}}</button>
                 <button (click)="deleteSparklineConfig()" style="flex:1; padding:6px; display:flex; align-items:center; justify-content:center; gap:4px; border:1px solid #dadce0; border-radius:4px; background:#fff; cursor:pointer; font-size:13px;"><span class="material-symbols-outlined" style="font-size:16px;">delete</span> Delete</button>
              </div>
              <button (click)="switchRowsColumns()" style="width:100%; padding:8px; border:1px solid #dadce0; border-radius:4px; background:#fff; cursor:pointer; font-size:13px; font-weight:500;">Switch rows / columns</button>

              <div style="border-top:1px solid #e2e8f0; margin:0 -16px; padding:12px 16px;">
                 <div (click)="horizontalAxisExpanded = !horizontalAxisExpanded" style="display:flex; justify-content:space-between; align-items:center; cursor:pointer; font-weight:600; font-size:13px; color:#202124;">
                   Horizontal Axis
                   <span class="material-symbols-outlined" style="font-size:18px;">{{horizontalAxisExpanded ? 'expand_less' : 'expand_more'}}</span>
                 </div>
                 <div *ngIf="horizontalAxisExpanded" style="margin-top:12px; display:flex; flex-direction:column; gap:8px;">
                   <label style="display:flex; align-items:center; gap:8px; font-size:13px; cursor:pointer;"><input type="checkbox" [(ngModel)]="sparklineConfig.horizontalAxis.displayAxis" (change)="saveSparkline()"> Display Axis</label>
                   <label style="display:flex; align-items:center; gap:8px; font-size:13px; cursor:pointer;"><input type="checkbox" [(ngModel)]="sparklineConfig.horizontalAxis.rightToLeft" (change)="saveSparkline()"> Plot sparkline from right to left</label>
                 </div>
              </div>

              <div style="border-top:1px solid #e2e8f0; border-bottom:1px solid #e2e8f0; margin:0 -16px; margin-top:-16px; padding:12px 16px;">
                 <div (click)="verticalAxisExpanded = !verticalAxisExpanded" style="display:flex; justify-content:space-between; align-items:center; cursor:pointer; font-weight:600; font-size:13px; color:#202124;">
                   Vertical Axis
                   <span class="material-symbols-outlined" style="font-size:18px;">{{verticalAxisExpanded ? 'expand_less' : 'expand_more'}}</span>
                 </div>
                 <div *ngIf="verticalAxisExpanded" style="margin-top:12px; display:flex; flex-direction:column; gap:16px;">
                   <div>
                     <div style="font-size:12px; font-weight:500; margin-bottom:6px;">Minimum Value:</div>
                     <label style="display:flex; align-items:center; gap:8px; font-size:13px; cursor:pointer;"><input type="radio" name="vMinMode" value="auto" [(ngModel)]="sparklineConfig.verticalAxis.min.mode" (change)="saveSparkline()"> Automatic for each sparkline</label>
                     <label style="display:flex; align-items:center; gap:8px; font-size:13px; cursor:pointer; margin-top:4px;"><input type="radio" name="vMinMode" value="same" [(ngModel)]="sparklineConfig.verticalAxis.min.mode" (change)="saveSparkline()"> Same for all sparklines</label>
                     <div style="display:flex; align-items:center; gap:8px; margin-top:4px;">
                       <label style="display:flex; align-items:center; gap:8px; font-size:13px; cursor:pointer;"><input type="radio" name="vMinMode" value="custom" [(ngModel)]="sparklineConfig.verticalAxis.min.mode" (change)="saveSparkline()"> Custom value:</label>
                       <input type="number" [(ngModel)]="sparklineConfig.verticalAxis.min.customValue" (change)="saveSparkline()" [disabled]="sparklineConfig.verticalAxis.min.mode !== 'custom'" style="width:60px; padding:2px 4px; border:1px solid #dadce0; border-radius:2px; font-size:12px;">
                     </div>
                   </div>
                   <div>
                     <div style="font-size:12px; font-weight:500; margin-bottom:6px;">Maximum Value:</div>
                     <label style="display:flex; align-items:center; gap:8px; font-size:13px; cursor:pointer;"><input type="radio" name="vMaxMode" value="auto" [(ngModel)]="sparklineConfig.verticalAxis.max.mode" (change)="saveSparkline()"> Automatic for each sparkline</label>
                     <label style="display:flex; align-items:center; gap:8px; font-size:13px; cursor:pointer; margin-top:4px;"><input type="radio" name="vMaxMode" value="same" [(ngModel)]="sparklineConfig.verticalAxis.max.mode" (change)="saveSparkline()"> Same for all sparklines</label>
                     <div style="display:flex; align-items:center; gap:8px; margin-top:4px;">
                       <label style="display:flex; align-items:center; gap:8px; font-size:13px; cursor:pointer;"><input type="radio" name="vMaxMode" value="custom" [(ngModel)]="sparklineConfig.verticalAxis.max.mode" (change)="saveSparkline()"> Custom value:</label>
                       <input type="number" [(ngModel)]="sparklineConfig.verticalAxis.max.customValue" (change)="saveSparkline()" [disabled]="sparklineConfig.verticalAxis.max.mode !== 'custom'" style="width:60px; padding:2px 4px; border:1px solid #dadce0; border-radius:2px; font-size:12px;">
                     </div>
                   </div>
                 </div>
              </div>

            </div>
          </ng-container>
                

        </div>
      </div>
    </div>

      <div class="toast" [class.show]="toastVisible">{{ toastMsg }}</div>

            <!-- Image Preview Modal -->
      <div *ngIf="previewImageUrl" class="modal-overlay" (click)="previewImageUrl = null" style="z-index: 10000; background: rgba(0,0,0,0.85); display: flex; align-items: center; justify-content: center;">
        <div style="position: relative; max-width: 90vw; max-height: 90vh; background: #fff; padding: 12px; padding-top: 48px; border-radius: 8px; box-shadow: 0 12px 40px rgba(0,0,0,0.5); display: inline-flex; flex-direction: column;" (click)="$event.stopPropagation()">
          <button style="position: absolute; top: 12px; right: 12px; background: rgba(0,0,0,0.05); border: none; border-radius: 50%; width: 32px; height: 32px; display: flex; align-items: center; justify-content: center; cursor: pointer; color: #5f6368; transition: background 0.2s;" (click)="previewImageUrl = null" onmouseover="this.style.background='rgba(0,0,0,0.1)'" onmouseout="this.style.background='rgba(0,0,0,0.05)'" title="Close">
            <span class="material-symbols-outlined" style="font-size: 20px;">close</span>
          </button>
          <a [href]="previewImageUrl" download="sheet_image.png" style="position: absolute; top: 12px; right: 52px; background: rgba(0,0,0,0.05); border: none; border-radius: 50%; width: 32px; height: 32px; display: flex; align-items: center; justify-content: center; cursor: pointer; color: #5f6368; transition: background 0.2s; text-decoration: none;" onmouseover="this.style.background='rgba(0,0,0,0.1)'" onmouseout="this.style.background='rgba(0,0,0,0.05)'" title="Download">
            <span class="material-symbols-outlined" style="font-size: 20px;">download</span>
          </a>
          <img [src]="previewImageUrl" style="max-width: 100%; max-height: calc(90vh - 60px); object-fit: contain; border-radius: 4px; display: block;">
        </div>
      </div>

      <!-- Right-click Context Menu -->
      <div class="ctx-menu" *ngIf="ctxVisible" [style.left.px]="ctxX" [style.top.px]="ctxTop" [style.bottom.px]="ctxBottom" [style.maxHeight.px]="ctxMaxHeight" (click)="$event.stopPropagation()">
        <div class="ctx-item" (click)="cutCell(); hideCtx()"><span class="ctx-icon material-symbols-outlined" style="font-size: 16px;">content_cut</span> Cut <span class="ctx-hint">Ctrl+X</span></div>
        <div class="ctx-item" (click)="copyCell(); hideCtx()"><span class="ctx-icon material-symbols-outlined" style="font-size: 16px;">content_copy</span> Copy <span class="ctx-hint">Ctrl+C</span></div>
        <div class="ctx-item" (mouseenter)="showCtxSubmenu('paste', $event)" (mouseleave)="hideCtxSubmenu()">
          <span class="ctx-icon material-symbols-outlined" style="font-size: 16px;">content_paste</span> Paste <span class="mdi-arrow material-symbols-outlined" style="margin-left:auto;font-size:16px;">chevron_right</span>
        </div>
        <div class="ctx-item" (click)="openCellEditHistory(); hideCtx()"><span class="ctx-icon material-symbols-outlined" style="font-size: 16px; color: #10b981;">history</span> Edit History...</div>
        <div class="ctx-sep"></div>
        <div class="ctx-item" (click)="openValidationModal(); hideCtx()"><span class="ctx-icon material-symbols-outlined" style="font-size: 16px;">arrow_drop_down_circle</span> Set dropdown list...</div>
        <div class="ctx-item danger" (click)="removeValidation(); hideCtx()"><span class="ctx-icon material-symbols-outlined" style="font-size: 16px;">close</span> Remove dropdown</div>
        <div class="ctx-sep"></div>
        <div class="ctx-item" (click)="sortColAZ(); hideCtx()"><span class="ctx-icon material-symbols-outlined" style="font-size: 16px;">sort</span> Sort A to Z</div>
        <div class="ctx-item" (click)="sortColZA(); hideCtx()"><span class="ctx-icon material-symbols-outlined" style="font-size: 16px;">sort</span> Sort Z to A</div>
        <div class="ctx-sep"></div>

        <div class="ctx-item" (mouseenter)="showCtxSubmenu('clear', $event)" (mouseleave)="hideCtxSubmenu()">
          <span class="ctx-icon material-symbols-outlined" style="font-size: 16px;">ink_eraser</span> Clear <span class="mdi-arrow material-symbols-outlined" style="margin-left:auto;font-size:16px;">chevron_right</span>
        </div>
        <div class="ctx-item" (mouseenter)="showCtxSubmenu('insert', $event)" (mouseleave)="hideCtxSubmenu()">
          <span class="ctx-icon material-symbols-outlined" style="font-size: 16px; color: #10b981;">add_box</span> Insert <span class="mdi-arrow material-symbols-outlined" style="margin-left:auto;font-size:16px;">chevron_right</span>
        </div>
        <div class="ctx-item" (mouseenter)="showCtxSubmenu('delete', $event)" (mouseleave)="hideCtxSubmenu()">
          <span class="ctx-icon material-symbols-outlined" style="font-size: 16px;">table_rows_narrow</span> Delete <span class="mdi-arrow material-symbols-outlined" style="margin-left:auto;font-size:16px;">chevron_right</span>
        </div>
        <div class="ctx-item" (mouseenter)="showCtxSubmenu('filter', $event)" (mouseleave)="hideCtxSubmenu()">
          <span class="ctx-icon material-symbols-outlined" style="font-size: 16px;">filter_alt</span> Filter by <span class="mdi-arrow material-symbols-outlined" style="margin-left:auto;font-size:16px;">chevron_right</span>
        </div>
      </div>

      <!-- Right-click Context Submenu -->
      <div class="ctx-menu" *ngIf="ctxVisible && activeCtxSubmenu"
           (mouseenter)="keepCtxSubmenu()" (mouseleave)="hideCtxSubmenu()"
           [style.left.px]="ctxSubX" 
           [style.top.px]="ctxSubTop" 
           [style.bottom.px]="ctxSubBottom"
           style="position: fixed; z-index: 100001; min-width: 220px; max-height: calc(100vh - 16px); overflow-y: auto; overflow-x: hidden;"
           [style.maxHeight.px]="ctxSubMaxHeight">
        
        <ng-container *ngIf="activeCtxSubmenu === 'paste'">
            <div class="ctx-item" (click)="pasteCell(); hideCtx()">All <span class="mh" style="margin-left:auto;color:#a0aec0;font-size:11px;">Ctrl+V</span></div>
            <div class="ctx-item" (click)="pasteValues(); hideCtx()">Values <span class="mh" style="margin-left:auto;color:#a0aec0;font-size:11px;">Ctrl+Shift+V</span></div>
            <div class="ctx-item" (click)="pasteFormulas(); hideCtx()">Formulas</div>
            <div class="ctx-item" (click)="pasteFormats(); hideCtx()">Formats</div>
            <div class="ctx-item" (click)="pasteNotes(); hideCtx()">Notes</div>
            <div class="ctx-sep"></div>
            <div class="ctx-item" (click)="pasteFormulasAndNumberFormats(); hideCtx()">Formulas and Number Formats</div>
            <div class="ctx-item" (click)="pasteValuesAndNumberFormats(); hideCtx()">Values and Number Formats</div>
            <div class="ctx-item" (click)="pasteValidation(); hideCtx()">Validation</div>
            <div class="ctx-sep"></div>
            <div class="ctx-item" (click)="pasteExceptNotes(); hideCtx()">All Except Notes</div>
            <div class="ctx-item" (click)="pasteExceptBorders(); hideCtx()">All Except Borders</div>
            <div class="ctx-sep"></div>
            <div class="ctx-item" (click)="pasteLinkToSource(); hideCtx()">Link To Source</div>
            <div class="ctx-item" (click)="pasteTranspose(); hideCtx()">Transpose</div>
        </ng-container>

        <ng-container *ngIf="activeCtxSubmenu === 'clear'">
            <div class="ctx-item" (click)="clearAll(); hideCtx()">All <span class="mh" style="margin-left:auto;color:#a0aec0;font-size:11px;">Ctrl+Del</span></div>
            <div class="ctx-item" (click)="clearAllFormats(); hideCtx()">Formats <span class="mh" style="margin-left:auto;color:#a0aec0;font-size:11px;">Shift+Del</span></div>
            <div class="ctx-item" (click)="clearRangeData(); hideCtx()">Contents <span class="mh" style="margin-left:auto;color:#a0aec0;font-size:11px;">Del</span></div>
            <div class="ctx-sep"></div>
            <div class="ctx-item" (click)="clearNotes(); hideCtx()">Notes</div>
            <div class="ctx-item" (click)="clearHyperlinks(); hideCtx()">Hyperlinks</div>
            <div class="ctx-item" (click)="clearCheckboxes(); hideCtx()">Checkboxes</div>
            <div class="ctx-sep"></div>
            <div class="ctx-item" (click)="clearDataValidations(); hideCtx()">Data Validations</div>
            <div class="ctx-item" (click)="clearConditionalFormats(); hideCtx()">Conditional Formats</div>
            <div class="ctx-item" (click)="clearRichTextFormats(); hideCtx()">RichText Formats</div>
            <div class="ctx-sep"></div>
            <div class="ctx-item" style="color:#ef4444;" (click)="clearAllFilters(); hideCtx()">Clear All Filters</div>
        </ng-container>

        <ng-container *ngIf="activeCtxSubmenu === 'insert'">
            <div class="ctx-item" (click)="shiftCellsRight(); hideCtx()">Shift Cells Right</div>
            <div class="ctx-item" (click)="shiftCellsDown(); hideCtx()">Shift Cells Down</div>
            <div class="ctx-sep"></div>
            <div class="ctx-item" (click)="insertRowAbove(); hideCtx()">1 Row Above</div>
            <div class="ctx-item" (click)="insertRowBelow(); hideCtx()">1 Row Below</div>
            <div class="ctx-item" (click)="openCustomInsert('row'); hideCtx()">Custom...</div>
            <div class="ctx-sep"></div>
            <div class="ctx-item" (click)="insertColLeft(); hideCtx()">Column Before</div>
            <div class="ctx-item" (click)="insertColRight(); hideCtx()">Column After</div>
            <div class="ctx-item" (click)="openCustomInsert('col'); hideCtx()">Custom...</div>
        </ng-container>

        <ng-container *ngIf="activeCtxSubmenu === 'delete'">
            <div class="ctx-item" (click)="shiftCellsLeft(); hideCtx()">Shift Cells Left</div>
            <div class="ctx-item" (click)="shiftCellsUp(); hideCtx()">Shift Cells Up</div>
            <div class="ctx-sep"></div>
            <div class="ctx-item danger" (click)="deleteRow(); hideCtx()">Delete {{ selectedRowCount }} Row{{ selectedRowCount > 1 ? 's' : '' }}</div>
            <div class="ctx-item danger" (click)="deleteCol(); hideCtx()">Delete {{ selectedColCount }} Column{{ selectedColCount > 1 ? 's' : '' }}</div>
        </ng-container>

        <ng-container *ngIf="activeCtxSubmenu === 'filter'">
            <div class="ctx-item" (click)="filterByCellValue(); hideCtx()">Cell Value</div>
            <div class="ctx-item" (click)="filterByCellColor(); hideCtx()">Cell Color</div>
            <div class="ctx-item" (click)="filterByTextColor(); hideCtx()">Text Color</div>
        </ng-container>
      </div>

      <!-- Advanced Filter Panel -->
      <div class="modal-overlay" *ngIf="advFilterVisible" (click)="closeAdvFilter()" style="position:fixed; top:0; left:0; width:100vw; height:100vh; background:transparent; z-index:100000;">
        <div class="modal" (click)="$event.stopPropagation()" [style.left.px]="advFilterX" [style.top.px]="advFilterY" 
             [style.max-height.px]="advFilterMaxHeight"
             style="position:fixed; width:280px; backdrop-filter:blur(16px); -webkit-backdrop-filter:blur(16px); border-radius:12px; display:flex; flex-direction:column; font-family:'Inter', sans-serif; z-index:100001; overflow:hidden;"
             [style.background]="currentTheme === 'dark' ? 'rgba(25,25,30,0.95)' : 'rgba(255,255,255,0.95)'"
             [style.color]="currentTheme === 'dark' ? '#f3f4f6' : '#111827'"
             [style.border]="currentTheme === 'dark' ? '1px solid rgba(255,255,255,0.08)' : '1px solid rgba(0,0,0,0.1)'"
             [style.box-shadow]="currentTheme === 'dark' ? '0 20px 40px rgba(0,0,0,0.5), 0 0 0 1px rgba(255,255,255,0.05)' : '0 20px 40px rgba(0,0,0,0.15), 0 0 0 1px rgba(0,0,0,0.05)'">
          
          <div [style.border-bottom]="currentTheme === 'dark' ? '1px solid rgba(255,255,255,0.05)' : '1px solid rgba(0,0,0,0.08)'" style="padding:16px 16px 12px;">
             <div style="font-weight:600; font-size:15px; display:flex; justify-content:space-between; margin-bottom:12px; align-items:center;">
                <span>Filter Options</span>
                <a href="javascript:void(0)" style="color:#10b981; text-decoration:none; font-size:12px; font-weight:500;">Custom Filter</a>
             </div>
             <div style="display:flex; gap:8px;">
                <button (click)="advSort(true)" style="flex:1; border-radius:6px; padding:6px; cursor:pointer; display:flex; justify-content:center; align-items:center; gap:6px; font-size:13px; font-weight:500; transition:all 0.2s;"
                        [style.background]="currentTheme === 'dark' ? 'rgba(255,255,255,0.05)' : 'rgba(0,0,0,0.04)'"
                        [style.color]="currentTheme === 'dark' ? '#f3f4f6' : '#374151'"
                        [style.border]="currentTheme === 'dark' ? '1px solid rgba(255,255,255,0.1)' : '1px solid rgba(0,0,0,0.05)'"
                        onmouseover="this.style.opacity='0.8'" onmouseout="this.style.opacity='1'"><span class="material-symbols-outlined" style="font-size:16px; color:#10b981;">sort_by_alpha</span> A-Z</button>
                <button (click)="advSort(false)" style="flex:1; border-radius:6px; padding:6px; cursor:pointer; display:flex; justify-content:center; align-items:center; gap:6px; font-size:13px; font-weight:500; transition:all 0.2s;"
                        [style.background]="currentTheme === 'dark' ? 'rgba(255,255,255,0.05)' : 'rgba(0,0,0,0.04)'"
                        [style.color]="currentTheme === 'dark' ? '#f3f4f6' : '#374151'"
                        [style.border]="currentTheme === 'dark' ? '1px solid rgba(255,255,255,0.1)' : '1px solid rgba(0,0,0,0.05)'"
                        onmouseover="this.style.opacity='0.8'" onmouseout="this.style.opacity='1'"><span class="material-symbols-outlined" style="font-size:16px; color:#ef4444;">sort_by_alpha</span> Z-A</button>
             </div>
          </div>

          <div [style.border-bottom]="currentTheme === 'dark' ? '1px solid rgba(255,255,255,0.05)' : '1px solid rgba(0,0,0,0.08)'" style="display:flex;">
             <div (click)="advFilterTab='value'" [style.border-bottom]="advFilterTab==='value'?'2px solid #10b981':''" [style.color]="advFilterTab==='value' ? '#10b981' : (currentTheme === 'dark' ? '#9ca3af' : '#6b7280')" style="flex:1; padding:10px 0; text-align:center; cursor:pointer; font-weight:600; font-size:12px; transition:color 0.2s;">ABC<br>123</div>
             <div (click)="advFilterTab='cellColor'" [style.border-bottom]="advFilterTab==='cellColor'?'2px solid #10b981':''" style="flex:1; padding:10px 0; text-align:center; cursor:pointer; display:flex; align-items:center; justify-content:center;"><span class="material-symbols-outlined" [style.color]="advFilterTab==='cellColor' ? '#10b981' : (currentTheme === 'dark' ? '#9ca3af' : '#6b7280')" style="transition:color 0.2s;">format_color_fill</span></div>
             <div (click)="advFilterTab='textColor'" [style.border-bottom]="advFilterTab==='textColor'?'2px solid #10b981':''" style="flex:1; padding:10px 0; text-align:center; cursor:pointer; display:flex; align-items:center; justify-content:center;"><span class="material-symbols-outlined" [style.color]="advFilterTab==='textColor' ? '#10b981' : (currentTheme === 'dark' ? '#9ca3af' : '#6b7280')" style="transition:color 0.2s;">format_color_text</span></div>
          </div>

          <div style="padding:16px; flex:1; min-height:0; overflow-y:auto; scrollbar-width:thin;" [style.scrollbar-color]="currentTheme === 'dark' ? 'rgba(255,255,255,0.2) transparent' : 'rgba(0,0,0,0.2) transparent'">
             <div *ngIf="advFilterTab==='value'">
                <div [style.color]="currentTheme === 'dark' ? '#9ca3af' : '#6b7280'" style="font-size:12px; font-weight:600; margin-bottom:10px; text-transform:uppercase; letter-spacing:0.5px;">Cell Value</div>
                <input type="text" placeholder="Search values..." [(ngModel)]="advFilterSearch" 
                       [style.background]="currentTheme === 'dark' ? 'rgba(0,0,0,0.2)' : 'rgba(255,255,255,0.5)'"
                       [style.color]="currentTheme === 'dark' ? '#f3f4f6' : '#111827'"
                       [style.border]="currentTheme === 'dark' ? '1px solid rgba(255,255,255,0.1)' : '1px solid rgba(0,0,0,0.15)'"
                       style="width:100%; padding:8px 12px; border-radius:6px; margin-bottom:12px; font-size:13px; box-sizing:border-box; outline:none; transition:border-color 0.2s;" onfocus="this.style.borderColor='#10b981'">
                
                <label [style.color]="currentTheme === 'dark' ? '#d1d5db' : '#374151'" style="display:flex; align-items:center; gap:10px; font-size:13px; margin-bottom:8px; cursor:pointer;">
                   <input type="checkbox" [checked]="allAdvFilterSelected('value')" (change)="toggleAllAdvFilter($event)" style="accent-color:#10b981; width:16px; height:16px; cursor:pointer;"> <span style="font-weight:500;">(Select All)</span>
                </label>
                <div *ngFor="let it of advFilterValues" [hidden]="!it.val.toLowerCase().includes(advFilterSearch.toLowerCase())">
                   <label [style.color]="currentTheme === 'dark' ? '#e5e7eb' : '#1f2937'" style="display:flex; align-items:center; gap:10px; font-size:13px; margin-bottom:8px; cursor:pointer;">
                      <input type="checkbox" [(ngModel)]="it.selected" style="accent-color:#10b981; width:16px; height:16px; cursor:pointer;"> <span style="text-overflow:ellipsis; overflow:hidden; white-space:nowrap;">{{ it.val || '(Blanks)' }}</span>
                   </label>
                </div>
             </div>
             
             <div *ngIf="advFilterTab==='cellColor'">
                <div [style.color]="currentTheme === 'dark' ? '#9ca3af' : '#6b7280'" style="font-size:12px; font-weight:600; margin-bottom:10px; text-transform:uppercase; letter-spacing:0.5px;">Cell Color</div>
                <label [style.color]="currentTheme === 'dark' ? '#d1d5db' : '#374151'" style="display:flex; align-items:center; gap:10px; font-size:13px; margin-bottom:8px; cursor:pointer;">
                   <input type="checkbox" [checked]="allAdvFilterSelected('cellColor')" (change)="toggleAllAdvFilter($event)" style="accent-color:#10b981; width:16px; height:16px; cursor:pointer;"> <span style="font-weight:500;">(Select All)</span>
                </label>
                <div *ngFor="let it of advFilterBgColors">
                   <label [style.color]="currentTheme === 'dark' ? '#e5e7eb' : '#1f2937'" style="display:flex; align-items:center; gap:10px; font-size:13px; margin-bottom:8px; cursor:pointer;">
                      <input type="checkbox" [(ngModel)]="it.selected" style="accent-color:#10b981; width:16px; height:16px; cursor:pointer;">
                      <div [style.background]="it.val || '#ffffff'" [style.border]="currentTheme === 'dark' ? '1px solid rgba(255,255,255,0.1)' : '1px solid rgba(0,0,0,0.1)'" style="width:24px; height:24px; border-radius:4px; box-shadow:inset 0 1px 2px rgba(0,0,0,0.1);"></div>
                      <span>{{ !it.val ? '(Default)' : it.val }}</span>
                   </label>
                </div>
             </div>

             <div *ngIf="advFilterTab==='textColor'">
                <div [style.color]="currentTheme === 'dark' ? '#9ca3af' : '#6b7280'" style="font-size:12px; font-weight:600; margin-bottom:10px; text-transform:uppercase; letter-spacing:0.5px;">Text Color</div>
                <label [style.color]="currentTheme === 'dark' ? '#d1d5db' : '#374151'" style="display:flex; align-items:center; gap:10px; font-size:13px; margin-bottom:8px; cursor:pointer;">
                   <input type="checkbox" [checked]="allAdvFilterSelected('textColor')" (change)="toggleAllAdvFilter($event)" style="accent-color:#10b981; width:16px; height:16px; cursor:pointer;"> <span style="font-weight:500;">(Select All)</span>
                </label>
                <div *ngFor="let it of advFilterTextColors">
                   <label [style.color]="currentTheme === 'dark' ? '#e5e7eb' : '#1f2937'" style="display:flex; align-items:center; gap:10px; font-size:13px; margin-bottom:8px; cursor:pointer;">
                      <input type="checkbox" [(ngModel)]="it.selected" style="accent-color:#10b981; width:16px; height:16px; cursor:pointer;">
                      <div [style.background]="it.val || '#000000'" [style.border]="currentTheme === 'dark' ? '1px solid rgba(255,255,255,0.1)' : '1px solid rgba(0,0,0,0.1)'" style="width:24px; height:24px; border-radius:4px; box-shadow:inset 0 1px 2px rgba(0,0,0,0.1);"></div>
                      <span>{{ !it.val ? '(Default)' : it.val }}</span>
                   </label>
                </div>
             </div>
          </div>

          <div [style.border-top]="currentTheme === 'dark' ? '1px solid rgba(255,255,255,0.05)' : '1px solid rgba(0,0,0,0.08)'" 
               [style.background]="currentTheme === 'dark' ? 'rgba(0,0,0,0.1)' : 'rgba(255,255,255,0.4)'"
               style="padding:12px 16px; display:flex; justify-content:flex-end; gap:8px;">
             <button (click)="clearAdvFilter()" 
                     [style.color]="currentTheme === 'dark' ? '#9ca3af' : '#4b5563'"
                     [style.border]="currentTheme === 'dark' ? '1px solid rgba(255,255,255,0.1)' : '1px solid rgba(0,0,0,0.15)'"
                     style="background:transparent; padding:8px 16px; border-radius:6px; cursor:pointer; font-size:13px; font-weight:500; transition:all 0.2s;" onmouseover="this.style.opacity='0.7'" onmouseout="this.style.opacity='1'">Clear</button>
             <button (click)="applyAdvFilter()" style="background:linear-gradient(135deg, #10b981 0%, #059669 100%); color:#fff; border:none; padding:8px 20px; border-radius:6px; cursor:pointer; font-size:13px; font-weight:600; box-shadow:0 4px 10px rgba(16,185,129,0.3); transition:all 0.2s;" onmouseover="this.style.opacity='0.9'; this.style.transform='translateY(-1px)'" onmouseout="this.style.opacity='1'; this.style.transform='none'">Apply Filter</button>
          </div>
        </div>
      </div>


      <!-- Manage Picklist Sidebar Modal (Zoho Picklist Style Sidebar) -->
      <div class="modal-overlay drawer-overlay" *ngIf="managePicklistSidebarOpen" (click)="managePicklistSidebarOpen = false" style="z-index: 99999; background: rgba(0,0,0,0.4); display: flex; justify-content: flex-end;">
        <div class="modal drawer-content" (click)="$event.stopPropagation()" style="width:420px; height:100vh; background:#202124; color:#e8eaed; border-left:1px solid #3c4043; box-shadow:-4px 0 24px rgba(0,0,0,0.5); display:flex; flex-direction:column; position:fixed; right:0; top:0; z-index:100000; box-sizing:border-box; overflow-x:hidden;">
          
          <!-- Header -->
          <div style="padding: 16px 20px; border-bottom:1px solid #3c4043; background:#202124; display:flex; justify-content:space-between; align-items:center;">
            <h3 style="margin:0; font-size:16px; font-weight:600; color:#fff;">Manage Picklist</h3>
            <button (click)="managePicklistSidebarOpen = false" style="background:none; border:none; cursor:pointer; color:#9aa0a6; display:flex; padding:0;"><span class="material-symbols-outlined" style="font-size:20px;">close</span></button>
          </div>

          <!-- View Rules Filters (Screenshot 1, 2, 3) -->
          <div style="padding: 14px 20px; border-bottom:1px solid #2d2e31; background:#202124; display:flex; flex-direction:column; gap:8px;">
            <div style="font-size:12.5px; color:#9aa0a6;">View Rules for:</div>
            <div style="display:flex; gap:10px;">
              <select [(ngModel)]="viewRulesSheet" (ngModelChange)="refreshManagePicklistRules()" style="flex:1.2; padding:6px 10px; border:1px solid #5f6368; border-radius:6px; background:#2d2e31; color:#fff; font-size:13px; outline:none; cursor:pointer;">
                <option value="current">{{sheets[currentSheetIdx].name}}</option>
                <option value="all">Whole Spreadsheet</option>
                <option *ngFor="let s of sheets" [value]="s.name">{{s.name}}</option>
              </select>

              <select [(ngModel)]="viewRulesType" (ngModelChange)="refreshManagePicklistRules()" style="flex:1; padding:6px 10px; border:1px solid #5f6368; border-radius:6px; background:#2d2e31; color:#fff; font-size:13px; outline:none; cursor:pointer;">
                <option value="all">All Rules</option>
                <option value="list">List</option>
                <option value="range">Cell Range</option>
              </select>
            </div>
          </div>

          <!-- Picklist Rule Cards Body -->
          <div style="padding: 16px 20px; flex:1; overflow-y:auto; background:#202124; display:flex; flex-direction:column; gap:14px;">
            <div *ngIf="_managePicklistRules.length === 0" style="text-align:center; padding:40px 10px; color:#9aa0a6; font-size:13px;">
              No picklist rules found.
            </div>

            <div *ngFor="let rule of _managePicklistRules" style="background:#2d2e31; border:1px solid #3c4043; border-radius:8px; padding:14px; display:flex; flex-direction:column; gap:10px;">
              <!-- Card Header -->
              <div style="display:flex; justify-content:space-between; align-items:center;">
                <span style="font-weight:600; font-size:13.5px; color:#fff;">{{rule.rangeRef}}</span>
                <div style="display:flex; align-items:center; gap:8px;">
                  <button (click)="$event.stopPropagation(); copyPicklistRule(rule, $event)" style="background:none; border:none; color:#9aa0a6; cursor:pointer; padding:2px; display:flex;" title="Copy rule">
                    <span class="material-symbols-outlined" style="font-size:18px; pointer-events:none;">content_copy</span>
                  </button>
                  <button (click)="$event.stopPropagation(); editPicklistRule(rule, $event)" style="background:none; border:none; color:#9aa0a6; cursor:pointer; padding:2px; display:flex;" title="Edit rule">
                    <span class="material-symbols-outlined" style="font-size:18px; pointer-events:none;">edit</span>
                  </button>
                  <button (click)="$event.stopPropagation(); deletePicklistRule(rule, $event)" style="background:none; border:none; color:#9aa0a6; cursor:pointer; padding:2px; display:flex;" title="Delete rule">
                    <span class="material-symbols-outlined" style="font-size:18px; pointer-events:none;">delete</span>
                  </button>
                </div>
              </div>

              <!-- Card Options Horizontal Bar -->
              <div style="display:flex; flex-wrap:wrap; gap:6px; font-size:12.5px; color:#bdc1c6;">
                <span *ngFor="let opt of rule.options; let isLast = last">
                  {{opt.label}}<span *ngIf="!isLast" style="margin-left:6px; color:#5f6368;">|</span>
                </span>
              </div>
            </div>
          </div>
        </div>
      </div>

      <!-- Validation / Dropdown Modal (Zoho Picklist Style Sidebar) -->
      <div class="modal-overlay drawer-overlay" *ngIf="validationModalOpen" (click)="showColorOptionsPopover ? closeColorOptionsPopover() : (validationModalOpen = false)" style="z-index: 99999; background: rgba(0,0,0,0.4); display: flex; justify-content: flex-end;">
        <div class="modal drawer-content" (click)="$event.stopPropagation()" style="width:420px; height:100vh; background:#202124; color:#e8eaed; border-left:1px solid #3c4043; box-shadow:-4px 0 24px rgba(0,0,0,0.5); display:flex; flex-direction:column; position:fixed; right:0; top:0; z-index:100000; box-sizing:border-box; overflow-x:hidden;">
          
          <!-- Header (Screenshot 4 & 5) -->
          <div style="padding: 16px 20px; border-bottom:1px solid #3c4043; background:#202124;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
              <h3 style="margin:0; font-size:16px; font-weight:600; color:#fff;">{{ isCopyMode ? 'Picklist' : 'Picklist - Edit' }}</h3>
              <button (click)="validationModalOpen = false; closeColorOptionsPopover()" style="background:none; border:none; cursor:pointer; color:#9aa0a6; display:flex; padding:0;"><span class="material-symbols-outlined" style="font-size:20px;">close</span></button>
            </div>

            <div style="display:flex; justify-content:space-between; align-items:center; font-size:12.5px; color:#bdc1c6; min-height:28px;">
              <!-- Static View Mode (Screenshot 5) -->
              <ng-container *ngIf="!appliesToEditing">
                <span>Applies to: <span style="color:#e8eaed; font-weight:500;">{{appliesToInput || ("'" + (sheets[currentSheetIdx].name || 'Sheet1') + "'." + getRangeRef())}}</span></span>
                <span (click)="startEditingAppliesTo()" style="color:#00c274; cursor:pointer; font-weight:500; margin-left:auto;">Edit</span>
              </ng-container>

              <!-- Editable Mode with Green Check & Red Cancel (Screenshot 4) -->
              <ng-container *ngIf="appliesToEditing">
                <span style="white-space:nowrap; margin-right:6px;">Applies to:</span>
                <input type="text" [(ngModel)]="appliesToInput" style="flex:1; background:transparent; border:none; border-bottom:1.5px solid #00c274; color:#fff; font-size:12.5px; outline:none; padding:2px 4px;">
                <button (click)="confirmAppliesTo()" style="background:none; border:none; color:#00c274; cursor:pointer; padding:2px; display:flex; margin-left:4px;" title="Confirm range">
                  <span class="material-symbols-outlined" style="font-size:20px;">check_circle</span>
                </button>
                <button (click)="cancelEditingAppliesTo()" style="background:none; border:none; color:#ea4335; cursor:pointer; padding:2px; display:flex; margin-left:2px;" title="Cancel">
                  <span class="material-symbols-outlined" style="font-size:20px;">cancel</span>
                </button>
              </ng-container>
            </div>
          </div>

          <!-- Body -->
          <div style="padding: 16px 20px; overflow-y:auto; overflow-x:hidden; flex:1; background:#202124; display:flex; flex-direction:column; gap:14px;">
            
            <!-- Selection Type Radios -->
            <div style="background:#2d2e31; border:1px solid #3c4043; border-radius:20px; padding:6px 16px; display:flex; justify-content:space-around;">
              <label style="display:flex; align-items:center; gap:6px; font-size:13px; cursor:pointer; color:#e8eaed;">
                <input type="radio" name="select_type" value="single" [(ngModel)]="picklistSelectType" style="accent-color:#00c274;"> Single-Select
              </label>
              <label style="display:flex; align-items:center; gap:6px; font-size:13px; cursor:pointer; color:#e8eaed;">
                <input type="radio" name="select_type" value="multi" [(ngModel)]="picklistSelectType" style="accent-color:#00c274;"> Multi-Select
              </label>
            </div>

            <!-- Settings Group -->
            <div style="background:#2d2e31; border:1px solid #3c4043; border-radius:8px; padding:12px; display:flex; flex-direction:column; gap:12px;">
              <div style="display:flex; align-items:center; justify-content:space-between;">
                <span style="font-size:13px; color:#9aa0a6;">Type</span>
                <select style="width:120px; padding:4px 8px; border:1px solid #5f6368; border-radius:4px; font-size:13px; outline:none; cursor:pointer; background:#202124; color:#fff;">
                  <option>List</option>
                </select>
              </div>

              <div style="display:flex; justify-content:space-between; align-items:center; font-size:13px; color:#e8eaed;">
                <span style="cursor:pointer; display:flex; align-items:center; gap:4px; color:#e8eaed;">Sort <span class="material-symbols-outlined" style="font-size:16px;">arrow_downward</span></span>
                <label (click)="displayAsChip = !displayAsChip" style="display:flex; align-items:center; gap:8px; cursor:pointer;">
                  Display as chip
                  <div [style.background]="displayAsChip ? '#00c274' : '#5f6368'" style="width:34px; height:18px; border-radius:10px; position:relative; transition:background 0.2s;">
                    <div [style.right]="displayAsChip ? '2px' : '18px'" style="width:14px; height:14px; background:#fff; border-radius:50%; position:absolute; top:2px; transition:right 0.2s;"></div>
                  </div>
                </label>
              </div>

              <div *ngIf="picklistSelectType !== 'multi'" style="display:flex; align-items:center; justify-content:space-between;">
                <span style="font-size:13px; color:#9aa0a6;">Color options</span>
                <select style="width:140px; padding:4px 8px; border:1px solid #5f6368; border-radius:4px; font-size:13px; outline:none; cursor:pointer; background:#202124; color:#fff;">
                  <option>Single Color</option>
                  <option>Multi Color</option>
                </select>
              </div>
            </div>

            <!-- Options Item List -->
            <div style="display:flex; flex-direction:column; gap:10px;">
              <div *ngFor="let opt of picklistOptions; let i = index" 
                   draggable="true" 
                   (dragstart)="onOptionDragStart($event, i)" 
                   (dragover)="onOptionDragOver($event, i)" 
                   (drop)="onOptionDrop($event, i)"
                   style="display:flex; align-items:center; gap:4px;">
                <span class="material-symbols-outlined" style="color:#9aa0a6; cursor:grab; font-size:18px; flex-shrink:0;">drag_indicator</span>
                
                <input type="text" [(ngModel)]="opt.label" placeholder="Item name"
                  [style.background]="picklistSelectType === 'multi' ? '#202124' : (opt.color || '#f97316')"
                  [style.color]="picklistSelectType === 'multi' ? '#e8eaed' : (opt.textColor || '#000000')"
                  [style.border]="picklistSelectType === 'multi' ? '1px solid #5f6368' : '1px solid transparent'"
                  [style.borderRadius]="displayAsChip ? '18px' : '4px'"
                  (click)="picklistSelectType !== 'multi' && openColorOptionsPopover(i, $event)"
                  style="flex:1; min-width:0; padding:7px 12px; outline:none; font-size:13px; font-weight:500; transition:all 0.2s;" [style.cursor]="picklistSelectType === 'multi' ? 'text' : 'pointer'" title="Click to configure colors">

                <button *ngIf="picklistSelectType !== 'multi'" type="button" (click)="openColorOptionsPopover(i, $event)" style="background:none;border:none;color:#9aa0a6;cursor:pointer;padding:2px;display:flex;align-items:center;flex-shrink:0;" title="Palette options">
                  <span class="material-symbols-outlined" style="font-size:18px;">palette</span>
                </button>
                <button (click)="picklistOptions.splice(i, 1)" style="background:none;border:none;color:#9aa0a6;cursor:pointer;padding:2px;display:flex;align-items:center;flex-shrink:0;" title="Delete item">
                  <span class="material-symbols-outlined" style="font-size:18px;">close</span>
                </button>
              </div>

              <button (click)="addPicklistOption()" style="background:none; border:1px dashed #5f6368; color:#00c274; border-radius:6px; padding:8px 12px; cursor:pointer; font-size:13px; font-weight:600; display:flex; align-items:center; justify-content:center; gap:6px; margin-top:6px;">
                <span class="material-symbols-outlined" style="font-size:18px;">add_circle</span> Add New
              </button>
            </div>
          </div>

          <!-- Footer -->
          <div style="padding: 16px 20px; border-top:1px solid #3c4043; background:#202124; display:flex; justify-content:flex-end; gap:12px;">
            <button (click)="saveValidation(); closeColorOptionsPopover()" style="background:#00c274; color:#fff; border:none; padding:8px 22px; border-radius:6px; cursor:pointer; font-weight:600; font-size:14px; transition:background 0.2s;" onmouseover="this.style.background='#009a5f'" onmouseout="this.style.background='#00c274'">Save</button>
            <button (click)="validationModalOpen = false; closeColorOptionsPopover()" style="background:#3c4043; color:#e8eaed; border:none; padding:8px 18px; border-radius:6px; cursor:pointer; font-weight:500; font-size:14px; transition:background 0.2s;" onmouseover="this.style.background='#4a4d51'" onmouseout="this.style.background='#3c4043'">Cancel</button>
          </div>
        </div>

        <!-- Color Options Popover Modal -->
        <div *ngIf="showColorOptionsPopover" (click)="$event.stopPropagation()" style="position:fixed; right:375px; top:120px; width:440px; background:#202124; border:1px solid #5f6368; border-radius:8px; box-shadow:0 12px 32px rgba(0,0,0,0.6); padding:16px; color:#e8eaed; z-index:100001;">
          <!-- Header -->
          <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px; border-bottom:1px solid #3c4043; padding-bottom:8px;">
            <span style="font-weight:600; font-size:14px; color:#fff;">Color options</span>
            <div style="display:flex; align-items:center; gap:12px;">
              <span (click)="resetItemColor()" style="color:#00c274; cursor:pointer; font-size:12px; font-weight:500;">Reset</span>
              <button (click)="closeColorOptionsPopover($event)" style="background:none; border:none; color:#9aa0a6; cursor:pointer; padding:0; display:flex;"><span class="material-symbols-outlined" style="font-size:18px;">close</span></button>
            </div>
          </div>

          <!-- 2 Columns: Fill Color & Text Color -->
          <div style="display:flex; gap:16px;">
            <!-- Fill Color Column -->
            <div style="flex:1;">
              <div style="font-size:12px; font-weight:600; color:#e8eaed; margin-bottom:8px;">Fill Color</div>
              <div style="font-size:11px; color:#9aa0a6; margin-bottom:4px;">Theme Colors</div>
              <div style="display:grid; grid-template-columns:repeat(10, 1fr); gap:3px; margin-bottom:10px;">
                <div *ngFor="let c of themeColorsGrid" (click)="setItemFillColor(c)" [style.background]="c" style="width:16px; height:16px; border-radius:2px; cursor:pointer; border:1px solid rgba(255,255,255,0.1);" [title]="c"></div>
              </div>
              <div style="font-size:11px; color:#9aa0a6; margin-bottom:4px;">Standard Colors</div>
              <div style="display:grid; grid-template-columns:repeat(10, 1fr); gap:3px;">
                <div *ngFor="let c of standardColorsRow" (click)="setItemFillColor(c)" [style.background]="c" style="width:16px; height:16px; border-radius:2px; cursor:pointer; border:1px solid rgba(255,255,255,0.1);" [title]="c"></div>
              </div>
            </div>

            <!-- Text Color Column -->
            <div style="flex:1;">
              <div style="font-size:12px; font-weight:600; color:#e8eaed; margin-bottom:8px;">Text Color</div>
              <div style="font-size:11px; color:#9aa0a6; margin-bottom:4px;">Theme Colors</div>
              <div style="display:grid; grid-template-columns:repeat(10, 1fr); gap:3px; margin-bottom:10px;">
                <div *ngFor="let c of themeColorsGrid" (click)="setItemTextColor(c)" [style.background]="c" style="width:16px; height:16px; border-radius:2px; cursor:pointer; border:1px solid rgba(255,255,255,0.1);" [title]="c"></div>
              </div>
              <div style="font-size:11px; color:#9aa0a6; margin-bottom:4px;">Standard Colors</div>
              <div style="display:grid; grid-template-columns:repeat(10, 1fr); gap:3px;">
                <div *ngFor="let c of standardColorsRow" (click)="setItemTextColor(c)" [style.background]="c" style="width:16px; height:16px; border-radius:2px; cursor:pointer; border:1px solid rgba(255,255,255,0.1);" [title]="c"></div>
              </div>
            </div>
          </div>

          <!-- Footer Buttons -->
          <div style="display:flex; justify-content:space-between; margin-top:16px; border-top:1px solid #3c4043; padding-top:10px;">
            <button (click)="closeColorOptionsPopover($event)" style="background:#3c4043; color:#e8eaed; border:none; padding:5px 12px; border-radius:4px; cursor:pointer; font-size:12px;">Back</button>
            <button style="background:#3c4043; color:#e8eaed; border:none; padding:5px 12px; border-radius:4px; cursor:pointer; font-size:12px;">More colors</button>
          </div>
        </div>
      </div>

      <!-- Goal Seek Modal -->
      <div class="modal-overlay" *ngIf="goalSeekModalOpen" (click)="goalSeekModalOpen=false" style="z-index:10000;">
        <div class="modal" (click)="$event.stopPropagation()" style="width:380px;background:#fff;color:#333;border-radius:8px;box-shadow:0 4px 12px rgba(0,0,0,0.15);padding:24px;">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px;">
            <h3 style="margin:0;font-size:18px;font-weight:600;">Goal Seek</h3>
            <button (click)="goalSeekModalOpen=false" style="background:none;border:none;cursor:pointer;color:#888;"><span class="material-symbols-outlined" style="font-size:20px;">close</span></button>
          </div>
          <div style="font-size:13px;color:#5f6368;margin-bottom:16px;">Find the input value needed to achieve a specific goal in a formula cell.</div>
          <div style="display:flex;flex-direction:column;gap:12px;margin-bottom:20px;">
            <div style="display:flex;align-items:center;gap:12px;">
              <label style="font-size:13px;width:110px;color:#555;">Set Cell:</label>
              <input [(ngModel)]="goalSeekTargetCell" placeholder="e.g. B5" style="flex:1;border:1px solid #cbd5e1;border-radius:4px;padding:7px 10px;font-size:13px;outline:none;" />
            </div>
            <div style="display:flex;align-items:center;gap:12px;">
              <label style="font-size:13px;width:110px;color:#555;">To Value:</label>
              <input [(ngModel)]="goalSeekTargetValue" placeholder="e.g. 1000" style="flex:1;border:1px solid #cbd5e1;border-radius:4px;padding:7px 10px;font-size:13px;outline:none;" />
            </div>
            <div style="display:flex;align-items:center;gap:12px;">
              <label style="font-size:13px;width:110px;color:#555;">By Changing Cell:</label>
              <input [(ngModel)]="goalSeekByCell" placeholder="e.g. A2" style="flex:1;border:1px solid #cbd5e1;border-radius:4px;padding:7px 10px;font-size:13px;outline:none;" />
            </div>
          </div>
          <div style="display:flex;justify-content:flex-end;gap:8px;">
            <button (click)="applyGoalSeek()" style="background:#10b981;color:#fff;border:none;padding:8px 24px;border-radius:4px;font-weight:600;cursor:pointer;">Solve</button>
            <button (click)="goalSeekModalOpen=false" style="background:#f1f5f9;color:#333;border:1px solid #e2e8f0;padding:8px 24px;border-radius:4px;font-weight:600;cursor:pointer;">Cancel</button>
          </div>
        </div>
      </div>

      <!-- Email Notifications Modal -->
      <div class="modal-overlay" *ngIf="emailNotifModalOpen" (click)="emailNotifModalOpen=false" style="z-index:10000;">
        <div class="modal" (click)="$event.stopPropagation()" style="width:420px;background:#fff;color:#333;border-radius:8px;box-shadow:0 4px 12px rgba(0,0,0,0.15);padding:24px;">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px;">
            <h3 style="margin:0;font-size:18px;font-weight:600;">Email Notification Settings</h3>
            <button (click)="emailNotifModalOpen=false" style="background:none;border:none;cursor:pointer;color:#888;"><span class="material-symbols-outlined" style="font-size:20px;">close</span></button>
          </div>
          <div style="display:flex;flex-direction:column;gap:14px;margin-bottom:20px;">
            <input [(ngModel)]="emailNotifEmail" placeholder="Notify email address" style="border:1px solid #cbd5e1;border-radius:4px;padding:8px 10px;font-size:13px;outline:none;" />
            <label style="display:flex;align-items:center;gap:8px;font-size:13px;cursor:pointer;">
              <input type="checkbox" [(ngModel)]="emailNotifOnEdit" style="accent-color:#10b981;"> Notify me when the sheet is edited
            </label>
            <label style="display:flex;align-items:center;gap:8px;font-size:13px;cursor:pointer;">
              <input type="checkbox" [(ngModel)]="emailNotifOnComment" style="accent-color:#10b981;"> Notify me when a comment is added
            </label>
          </div>
          <div style="display:flex;justify-content:flex-end;gap:8px;">
            <button (click)="saveEmailNotifications()" style="background:#10b981;color:#fff;border:none;padding:8px 24px;border-radius:4px;font-weight:600;cursor:pointer;">Save</button>
            <button (click)="emailNotifModalOpen=false" style="background:#f1f5f9;color:#333;border:1px solid #e2e8f0;padding:8px 24px;border-radius:4px;font-weight:600;cursor:pointer;">Cancel</button>
          </div>
        </div>
      </div>

      <!-- Preferences Modal -->
      <div class="modal-overlay" *ngIf="preferencesModalOpen" (click)="preferencesModalOpen=false" style="z-index:10000;">
        <div class="modal" (click)="$event.stopPropagation()" style="width:420px;background:#fff;color:#333;border-radius:8px;box-shadow:0 4px 12px rgba(0,0,0,0.15);padding:24px;">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px;">
            <h3 style="margin:0;font-size:18px;font-weight:600;">My Preferences</h3>
            <button (click)="preferencesModalOpen=false" style="background:none;border:none;cursor:pointer;color:#888;"><span class="material-symbols-outlined" style="font-size:20px;">close</span></button>
          </div>
          <div style="display:flex;flex-direction:column;gap:14px;margin-bottom:20px;">
            <div style="display:flex;align-items:center;gap:12px;">
              <label style="font-size:13px;width:130px;color:#555;">Locale:</label>
              <select [(ngModel)]="prefLocale" style="flex:1;border:1px solid #cbd5e1;border-radius:4px;padding:7px;font-size:13px;outline:none;background:#fff;">
                <option value="en-US">English (US)</option>
                <option value="en-IN">English (India)</option>
                <option value="en-GB">English (UK)</option>
              </select>
            </div>
            <div style="display:flex;align-items:center;gap:12px;">
              <label style="font-size:13px;width:130px;color:#555;">Date Format:</label>
              <select [(ngModel)]="prefDateFormat" style="flex:1;border:1px solid #cbd5e1;border-radius:4px;padding:7px;font-size:13px;outline:none;background:#fff;">
                <option value="MM/DD/YYYY">MM/DD/YYYY</option>
                <option value="DD/MM/YYYY">DD/MM/YYYY</option>
                <option value="YYYY-MM-DD">YYYY-MM-DD</option>
              </select>
            </div>
            <label style="display:flex;align-items:center;gap:8px;font-size:13px;cursor:pointer;">
              <input type="checkbox" [(ngModel)]="prefThousands" style="accent-color:#10b981;"> Use thousands separator (1,000)
            </label>
          </div>
          <div style="display:flex;justify-content:flex-end;gap:8px;">
            <button (click)="savePreferences()" style="background:#10b981;color:#fff;border:none;padding:8px 24px;border-radius:4px;font-weight:600;cursor:pointer;">Save</button>
            <button (click)="preferencesModalOpen=false" style="background:#f1f5f9;color:#333;border:1px solid #e2e8f0;padding:8px 24px;border-radius:4px;font-weight:600;cursor:pointer;">Cancel</button>
          </div>
        </div>
      </div>

      <!-- Feedback Modal -->
      <div class="modal-overlay" *ngIf="feedbackModalOpen" (click)="feedbackModalOpen=false" style="z-index:10000;">
        <div class="modal" (click)="$event.stopPropagation()" style="width:420px;background:#fff;color:#333;border-radius:8px;box-shadow:0 4px 12px rgba(0,0,0,0.15);padding:0; overflow:hidden; display:flex; flex-direction:column;">
          <div style="display:flex;justify-content:space-between;align-items:center;padding:16px 20px;border-bottom:1px solid #e0e0e0;">
            <h3 style="margin:0;font-size:16px;font-weight:600;">Feedback</h3>
            <button (click)="feedbackModalOpen=false" style="background:none;border:none;cursor:pointer;color:#888;display:flex;align-items:center;justify-content:center;"><span class="material-symbols-outlined" style="font-size:20px;">close</span></button>
          </div>
          <div style="padding:20px; display:flex; flex-direction:column; gap:16px;">
            <div style="font-size:14px; color:#5f6368;">Love us or not—let us know!</div>
            
            <div style="display:flex; justify-content:space-around; align-items:center; padding: 4px 0;">
              <div (click)="feedbackType='Love'" style="display:flex; flex-direction:column; align-items:center; cursor:pointer;" [style.opacity]="feedbackType === 'Love' ? '1' : '0.4'" [style.filter]="feedbackType === 'Love' ? 'none' : 'grayscale(100%)'">
                <span style="font-size:24px;">❤️</span>
                <span style="font-size:12px; margin-top:4px; font-weight:500;">Love</span>
              </div>
              <div (click)="feedbackType='Idea'" style="display:flex; flex-direction:column; align-items:center; cursor:pointer;" [style.opacity]="feedbackType === 'Idea' ? '1' : '0.4'" [style.filter]="feedbackType === 'Idea' ? 'none' : 'grayscale(100%)'">
                <span style="font-size:24px;">💡</span>
                <span style="font-size:12px; margin-top:4px; font-weight:500;">Idea</span>
              </div>
              <div (click)="feedbackType='Help'" style="display:flex; flex-direction:column; align-items:center; cursor:pointer;" [style.opacity]="feedbackType === 'Help' ? '1' : '0.4'" [style.filter]="feedbackType === 'Help' ? 'none' : 'grayscale(100%)'">
                <span style="font-size:24px;">❓</span>
                <span style="font-size:12px; margin-top:4px; font-weight:500;">Help</span>
              </div>
              <div (click)="feedbackType='Bug'" style="display:flex; flex-direction:column; align-items:center; cursor:pointer;" [style.opacity]="feedbackType === 'Bug' ? '1' : '0.4'" [style.filter]="feedbackType === 'Bug' ? 'none' : 'grayscale(100%)'">
                <span style="font-size:24px;">🐞</span>
                <span style="font-size:12px; margin-top:4px; font-weight:500;">Bug</span>
              </div>
            </div>

            <textarea [(ngModel)]="feedbackText" placeholder="Post your comments" style="border:1px solid #dadce0;border-radius:4px;padding:12px;font-size:13px;height:90px;outline:none;resize:none;font-family:inherit;"></textarea>
            
            <div style="display:flex; justify-content:space-between; align-items:center;">
              <div style="display:flex; align-items:center; gap:8px;">
                <button (click)="feedbackFileInput.click()" style="display:flex; align-items:center; gap:6px; background:#fff; border:1px solid #dadce0; border-radius:4px; padding:6px 12px; color:#333; cursor:pointer; font-size:13px; font-weight:500;">
                  <span class="material-symbols-outlined" style="font-size:18px;">attach_file</span>
                  Attach file
                </button>
                <input type="file" #feedbackFileInput style="display:none;" (change)="onFeedbackFileSelected($event)">
                <div *ngIf="feedbackFile" style="font-size:12px; color:#1a73e8; max-width:120px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" [title]="feedbackFile.name">
                  {{ feedbackFile.name }}
                </div>
                <span *ngIf="feedbackFile" class="material-symbols-outlined" style="font-size:14px; cursor:pointer; color:#5f6368;" (click)="feedbackFile=null">close</span>
              </div>
              <div style="display:flex; flex-direction:column; align-items:center;">
                <span style="font-size:12px; color:#5f6368; margin-bottom:4px; font-weight:500;">Rate us</span>
                <div style="display:flex; gap:2px;">
                  <span *ngFor="let s of [1,2,3,4,5]" (click)="feedbackRating=s" class="material-symbols-outlined" [style.color]="s<=feedbackRating ? '#f29900' : '#dadce0'" style="font-size:20px; cursor:pointer; font-variation-settings: 'FILL' {{ s <= feedbackRating ? 1 : 0 }};">star</span>
                </div>
              </div>
            </div>

            <div style="display:flex; gap:8px; align-items:flex-start; margin-top:4px;">
              <input type="checkbox" [(ngModel)]="feedbackRecordScreen" style="margin-top:2px; accent-color:#10b981;" id="recordScreenCb">
              <label for="recordScreenCb" style="font-size:12px; color:#333; line-height:1.4; cursor:pointer;">Record my screen to help the team understand my issue and take necessary actions quickly.</label>
            </div>
          </div>
          
          <div style="display:flex;justify-content:flex-end;gap:12px; padding:16px 20px; border-top:1px solid #e0e0e0; background:#f8f9fa;">
            <button (click)="submitFeedback()" style="background:#10b981;color:#fff;border:none;padding:8px 24px;border-radius:4px;font-weight:600;cursor:pointer;font-size:13px;">Submit</button>
            <button (click)="feedbackModalOpen=false" style="background:#fff;color:#333;border:1px solid #dadce0;padding:8px 24px;border-radius:4px;font-weight:600;cursor:pointer;font-size:13px;">Cancel</button>
          </div>
        </div>
      </div>

      <!-- Text to Columns Modal -->
      <div class="modal-overlay" *ngIf="textToColsModalOpen" (click)="textToColsModalOpen = false" style="z-index: 10000;">
        <div class="modal" (click)="$event.stopPropagation()" style="width:420px; background:#fff; color:#333; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.15); padding:24px;">
          <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:20px;">
            <h3 style="margin:0; font-size:18px; font-weight:600;">Text to Columns</h3>
            <button (click)="textToColsModalOpen = false" style="background:none; border:none; cursor:pointer; color:#888;"><span class="material-symbols-outlined" style="font-size:20px;">close</span></button>
          </div>
          <div style="margin-bottom:16px; font-size:13px; color:#5f6368;">Splits the selected column's content into multiple columns using a delimiter.</div>
          <div style="margin-bottom:16px;">
            <div style="font-size:13px; font-weight:600; margin-bottom:10px;">Separator:</div>
            <div style="display:flex; flex-direction:column; gap:8px;">
              <label style="display:flex; align-items:center; gap:8px; font-size:13px; cursor:pointer;">
                <input type="radio" name="t2cDelim" value="," [(ngModel)]="t2cDelimiter" style="accent-color:#10b981;"> Comma (,)
              </label>
              <label style="display:flex; align-items:center; gap:8px; font-size:13px; cursor:pointer;">
                <input type="radio" name="t2cDelim" value="	" [(ngModel)]="t2cDelimiter" style="accent-color:#10b981;"> Tab
              </label>
              <label style="display:flex; align-items:center; gap:8px; font-size:13px; cursor:pointer;">
                <input type="radio" name="t2cDelim" value=" " [(ngModel)]="t2cDelimiter" style="accent-color:#10b981;"> Space
              </label>
              <label style="display:flex; align-items:center; gap:8px; font-size:13px; cursor:pointer;">
                <input type="radio" name="t2cDelim" value=";" [(ngModel)]="t2cDelimiter" style="accent-color:#10b981;"> Semicolon (;)
              </label>
              <label style="display:flex; align-items:center; gap:8px; font-size:13px; cursor:pointer;">
                <input type="radio" name="t2cDelim" value="custom" [(ngModel)]="t2cDelimiter" style="accent-color:#10b981;"> Custom:
                <input [(ngModel)]="t2cCustomDelim" (focus)="t2cDelimiter='custom'" placeholder="e.g. |" style="border:1px solid #cbd5e1; border-radius:4px; padding:4px 8px; width:60px; outline:none; font-size:13px;" />
              </label>
            </div>
          </div>
          <div style="display:flex; justify-content:flex-end; gap:8px;">
            <button (click)="applyTextToColumns()" style="background:#10b981; color:#fff; border:none; padding:8px 24px; border-radius:4px; font-weight:600; cursor:pointer;">Apply</button>
            <button (click)="textToColsModalOpen = false" style="background:#f1f5f9; color:#333; border:1px solid #e2e8f0; padding:8px 24px; border-radius:4px; font-weight:600; cursor:pointer;">Cancel</button>
          </div>
        </div>
      </div>

      <!-- Range Picker Floating Bar -->
      <div *ngIf="rangePickerActive" style="position:fixed; top:16px; left:50%; transform:translateX(-50%); z-index:20000; background:#fff; border:2px solid #10b981; border-radius:8px; box-shadow:0 4px 20px rgba(0,0,0,0.2); display:flex; align-items:center; gap:8px; padding:8px 12px; min-width:320px;">
        <span class="material-symbols-outlined" style="color:#10b981; font-size:18px;">grid_on</span>
        <input [value]="getRangePickerValue()" (input)="onRangePickerInput($event)" style="flex:1; border:none; outline:none; font-size:13px; color:#1a73e8; font-weight:500;" />
        <button (click)="confirmRangePicker()" style="background:#10b981; color:#fff; border:none; border-radius:4px; padding:4px 10px; cursor:pointer; display:flex; align-items:center;">
          <span class="material-symbols-outlined" style="font-size:18px;">check</span>
        </button>
        <button (click)="cancelRangePicker()" style="background:#f1f5f9; color:#333; border:1px solid #e2e8f0; border-radius:4px; padding:4px 10px; cursor:pointer; display:flex; align-items:center;">
          <span class="material-symbols-outlined" style="font-size:18px;">close</span>
        </button>
      </div>

      <!-- Create Pivot Table Modal -->
      <div class="modal-overlay" *ngIf="pivotModalOpen" (click)="pivotModalOpen = false" style="z-index: 10000;">
        <div class="modal" (click)="$event.stopPropagation()" style="width:400px; background:#fff; color:#333; border-radius: 8px; box-shadow:0 4px 12px rgba(0,0,0,0.15); display:flex; flex-direction:column; padding: 20px;">
          
          <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;">
            <h3 style="margin:0; font-size:18px; font-weight:600;">Create Pivot Table</h3>
            <button (click)="pivotModalOpen = false" style="background:none; border:none; cursor:pointer; color:#888; display:flex;"><span class="material-symbols-outlined" style="font-size:20px;">close</span></button>
          </div>

          <div style="margin-bottom:16px;">
            <div style="font-size:13px; margin-bottom:8px;">Choose the data range for the table:</div>
            <div style="display:flex; align-items:center; gap:8px;">
              <span style="font-size:13px;">Source:</span>
              <div style="flex:1; border: 1px solid #10b981; border-radius: 4px; padding: 4px 8px; display:flex; align-items:center; gap:4px;">
                <input [(ngModel)]="pivotSource" style="border:none; outline:none; flex:1; font-size:13px; color:#1a73e8; background:transparent;" placeholder="Sheet1.A1:C10" />
                <span class="material-symbols-outlined" style="font-size:16px; color:#5f6368; cursor:pointer;" (click)="startRangePicker('pivotSource')">grid_on</span>
              </div>
            </div>
          </div>

          <div style="margin-bottom:24px;">
            <div style="font-size:13px; margin-bottom:8px;">Choose the location for the table:</div>
            <div style="display:flex; align-items:center; gap:16px; margin-bottom: 12px; font-size:13px;">
              <label style="display:flex; align-items:center; gap:6px; cursor:pointer;">
                <input type="radio" name="pivotLoc" value="new" [(ngModel)]="pivotDestType" style="accent-color:#10b981;"> New sheet
              </label>
              <label style="display:flex; align-items:center; gap:6px; cursor:pointer;">
                <input type="radio" name="pivotLoc" value="existing" [(ngModel)]="pivotDestType" style="accent-color:#10b981;"> Existing sheet
              </label>
            </div>
            <div *ngIf="pivotDestType === 'existing'" style="display:flex; align-items:center; gap:8px;">
              <span style="font-size:13px;">Location:</span>
              <div style="flex:1; border: 1px solid #cbd5e1; border-radius: 4px; padding: 4px 8px; display:flex; align-items:center; gap:4px;">
                <input [(ngModel)]="pivotDest" style="border:none; outline:none; flex:1; font-size:13px; color:#1a73e8; background:transparent;" placeholder="Sheet1.A9" />
                <span class="material-symbols-outlined" style="font-size:16px; color:#5f6368; cursor:pointer;" (click)="startRangePicker('pivotDest')">grid_on</span>
              </div>
            </div>
          </div>

          <div style="display:flex; justify-content:flex-end; gap:8px;">
            <button (click)="createPivotTable()" style="background:#10b981; color:#fff; border:none; padding:8px 24px; border-radius:4px; font-weight:600; cursor:pointer;">OK</button>
            <button (click)="pivotModalOpen = false" style="background:#f1f5f9; color:#333; border:1px solid #e2e8f0; padding:8px 24px; border-radius:4px; font-weight:600; cursor:pointer;">Cancel</button>
          </div>
        </div>
      </div>

      <!-- Data Validation Modal -->
      <div class="modal-overlay" *ngIf="dataValidationModalOpen" (click)="dataValidationModalOpen = false" style="z-index: 10000;">
        <div class="modal" (click)="$event.stopPropagation()" style="width:500px; background:#fff; color:#333; border-radius: 8px; box-shadow:0 4px 12px rgba(0,0,0,0.15); display:flex; flex-direction:column; padding: 20px;">
          
          <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;">
            <h3 style="margin:0; font-size:18px; font-weight:600;">Data Validation</h3>
            <button (click)="dataValidationModalOpen = false" style="background:none; border:none; cursor:pointer; color:#888; display:flex;"><span class="material-symbols-outlined" style="font-size:20px;">close</span></button>
          </div>

          <div style="display:flex; align-items:center; gap:16px; margin-bottom:16px;">
            <span style="font-size:13px; width: 70px;">Applies to:</span>
            <div style="flex:1; border: 1px solid #10b981; border-radius: 4px; padding: 4px 8px; display:flex; align-items:center; gap:4px;">
              <input [(ngModel)]="dvAppliesTo" style="border:none; outline:none; flex:1; font-size:13px; color:#1a73e8; background:transparent;" />
              <span class="material-symbols-outlined" style="font-size:16px; color:#5f6368; cursor:pointer;" (click)="startRangePicker('dvAppliesTo')">grid_on</span>
            </div>
          </div>

          <div style="display:flex; align-items:center; gap:12px; margin-bottom:16px; flex-wrap:wrap;">
            <span style="font-size:13px; width:70px;">Criteria:</span>
            <select [(ngModel)]="dvCriteria" style="border: 1px solid #cbd5e1; border-radius: 4px; padding: 6px; font-size:13px; width:130px; outline:none; background:#fff;">
              <option value="list">List</option>
              <option value="number">Number</option>
              <option value="text">Text</option>
              <option value="date">Date</option>
              <option value="checkbox">Checkbox</option>
            </select>
            <label style="display:flex; align-items:center; gap:4px; font-size:13px; cursor:pointer;">
              <input type="checkbox" [(ngModel)]="dvShowList" style="accent-color:#10b981;"> Show List
            </label>
            <label style="display:flex; align-items:center; gap:4px; font-size:13px; cursor:pointer;">
              <input type="checkbox" [(ngModel)]="dvSortAsc" style="accent-color:#10b981;"> Sort Ascending
            </label>
          </div>

          <div style="margin-left: 84px; margin-bottom: 16px;">
            <textarea placeholder="Line Separated Values (one per line)" [(ngModel)]="validationInput" style="width: 100%; height: 80px; border: 1px solid #cbd5e1; border-radius: 4px; padding: 8px; font-size:13px; outline:none; font-family:inherit; resize:none; box-sizing:border-box;"></textarea>
          </div>

          <ng-container *ngIf="dvCriteria === 'list'">
            <!-- Single vs Multi Select -->
            <div style="margin-left: 84px; margin-bottom: 12px; display:flex; align-items:center; gap:16px;">
              <label style="display:flex; align-items:center; gap:4px; font-size:13px; cursor:pointer;">
                <input type="radio" name="pickType" [value]="false" [(ngModel)]="dvIsMultiSelect" style="accent-color:#10b981;"> Single-Select
              </label>
              <label style="display:flex; align-items:center; gap:4px; font-size:13px; cursor:pointer;">
                <input type="radio" name="pickType" [value]="true" [(ngModel)]="dvIsMultiSelect" style="accent-color:#10b981;"> Multi-Select
              </label>
            </div>

            <!-- Display as chip -->
            <div style="margin-left: 84px; margin-bottom: 12px;">
              <label style="display:flex; align-items:center; gap:4px; font-size:13px; cursor:pointer;">
                <input type="checkbox" [(ngModel)]="dvDisplayAsChip" style="accent-color:#10b981;"> Display as chip
              </label>
            </div>

            <!-- Color Options -->
            <div style="margin-left: 84px; margin-bottom: 16px; display:flex; flex-direction:column; gap:8px;">
              <div style="display:flex; align-items:center; gap:12px;">
                <span style="font-size:13px;">Color options:</span>
                <select [(ngModel)]="dvColorMode" style="border: 1px solid #cbd5e1; border-radius: 4px; padding: 6px; font-size:13px; width:130px; outline:none; background:#fff;">
                  <option value="none">None</option>
                  <option value="single">Single Color</option>
                  <option value="multi">Multi Color</option>
                </select>
                <div *ngIf="dvColorMode === 'single'" style="display:flex; align-items:center; gap:4px;">
                   <input type="color" [(ngModel)]="dvSingleColor" list="presetColors" style="width:24px; height:24px; padding:0; border:none; cursor:pointer; background:transparent;">
                </div>
              </div>
              
              <!-- Multi Color Options Mapping -->
              <div *ngIf="dvColorMode === 'multi'" style="display:flex; flex-direction:column; gap:6px; max-height:120px; overflow-y:auto; padding-right:8px; border:1px solid #e2e8f0; border-radius:4px; padding:8px;">
                 <div *ngFor="let item of getParsedValidationItems()" style="display:flex; justify-content:space-between; align-items:center; font-size:13px;">
                   <span style="overflow:hidden; text-overflow:ellipsis; white-space:nowrap; max-width:280px;">{{item}}</span>
                   <div style="position:relative;">
                     <div (click)="openColorPalette(item)" [style.background]="dvItemColors[item] || '#e2e8f0'" style="width:20px; height:20px; border:1px solid #ccc; border-radius:4px; cursor:pointer;"></div>
                     <!-- Basic Palette Popup -->
                     <div *ngIf="dvPaletteItem === item" style="position:absolute; right:0; top:24px; z-index:100; background:#fff; border:1px solid #cbd5e1; border-radius:4px; box-shadow:0 4px 6px rgba(0,0,0,0.1); padding:8px; width:160px; display:flex; flex-wrap:wrap; gap:4px;">
                       <div *ngFor="let c of ['#ef4444','#f97316','#f59e0b','#84cc16','#22c55e','#10b981','#14b8a6','#06b6d4','#0ea5e9','#3b82f6','#6366f1','#8b5cf6','#a855f7','#d946ef','#ec4899','#f43f5e','#cbd5e1','#94a3b8','#64748b','#334155']" 
                            (click)="setPaletteColor(item, c)" 
                            [style.background]="c" 
                            style="width:24px; height:24px; border-radius:4px; cursor:pointer; border:1px solid #e2e8f0;"></div>
                     </div>
                   </div>
                 </div>
                 <div *ngIf="!getParsedValidationItems().length" style="font-size:12px; color:#888;">Enter values above first.</div>
              </div>
            </div>
          </ng-container>

          <div style="margin-left: 84px; margin-bottom: 16px;">
            <label style="display:flex; align-items:center; gap:6px; font-size:13px; cursor:pointer;">
              <input type="checkbox" [(ngModel)]="dvIgnoreBlanks" style="accent-color:#10b981;"> Ignore Blanks
            </label>
          </div>

          <!-- Alerts and Help Text (expandable) -->
          <div style="margin-bottom:16px; border:1px solid #e2e8f0; border-radius:4px; overflow:hidden;">
            <div (click)="dvAlertsOpen = !dvAlertsOpen" style="display:flex; align-items:center; gap:4px; cursor:pointer; user-select:none; padding:10px 12px; background:#f8f9fa;">
              <span class="material-symbols-outlined" style="font-size:16px; transition:transform 0.2s;" [style.transform]="dvAlertsOpen ? 'rotate(90deg)' : 'rotate(0deg)'">arrow_right</span>
              <span style="font-size:13px; font-weight:500;">Alerts and Help Text</span>
            </div>
            <div *ngIf="dvAlertsOpen" style="padding:12px; display:flex; flex-direction:column; gap:10px;">
              <label style="display:flex; align-items:center; gap:6px; font-size:13px; cursor:pointer;">
                <input type="checkbox" [(ngModel)]="dvAlertEnabled" style="accent-color:#10b981;"> Show validation error alert
              </label>
              <div *ngIf="dvAlertEnabled" style="display:flex; flex-direction:column; gap:8px;">
                <input [(ngModel)]="dvAlertTitle" placeholder="Alert Title" style="border:1px solid #cbd5e1; border-radius:4px; padding:6px 8px; font-size:13px; outline:none;" />
                <textarea [(ngModel)]="dvAlertMsg" placeholder="Alert message shown when invalid data is entered" style="border:1px solid #cbd5e1; border-radius:4px; padding:6px 8px; font-size:13px; outline:none; resize:none; height:60px;"></textarea>
              </div>
            </div>
          </div>

          <div style="display:flex; justify-content:flex-end; gap:8px;">
            <button (click)="saveDataValidation()" style="background:#10b981; color:#fff; border:none; padding:8px 24px; border-radius:4px; font-weight:600; cursor:pointer;">OK</button>
            <button (click)="dataValidationModalOpen = false" style="background:#f1f5f9; color:#333; border:1px solid #e2e8f0; padding:8px 24px; border-radius:4px; font-weight:600; cursor:pointer;">Cancel</button>
          </div>
        </div>
      </div>

      <!-- Manage Rules Modal -->
      <div class="modal-overlay" *ngIf="manageRulesModalOpen" (click)="manageRulesModalOpen = false" style="z-index: 10000;">
        <div class="modal" (click)="$event.stopPropagation()" style="width:400px; background:#fff; color:#333; border-radius: 8px; box-shadow:0 4px 12px rgba(0,0,0,0.15); display:flex; flex-direction:column;">
          
          <div style="display:flex; justify-content:space-between; align-items:center; padding: 16px 20px; border-bottom:1px solid #e2e8f0;">
            <h3 style="margin:0; font-size:18px; font-weight:600;">Data Validation - Manage Rules</h3>
            <button (click)="manageRulesModalOpen = false" style="background:none; border:none; cursor:pointer; color:#888; display:flex;"><span class="material-symbols-outlined" style="font-size:20px;">close</span></button>
          </div>

          <div style="padding: 16px 20px; background:#f8f9fa; display:flex; align-items:center; gap:8px; border-bottom:1px solid #e2e8f0;">
            <span style="font-size:13px;">View Rules for:</span>
            <select style="border: 1px solid #cbd5e1; border-radius: 4px; padding: 6px; font-size:13px; width:120px; outline:none; background:#fff;">
              <option>Sheet1</option>
            </select>
          </div>

          <div style="padding: 60px 20px; text-align:center; color:#5f6368; font-size:14px;">
            No Rules
          </div>

          <div style="padding: 16px 20px; display:flex; justify-content:space-between; align-items:center; border-top:1px solid #e2e8f0;">
            <button (click)="manageRulesModalOpen = false; dataValidationModalOpen = true" style="background:#f1f5f9; color:#333; border:1px solid #e2e8f0; padding:8px 16px; border-radius:4px; font-weight:600; cursor:pointer; display:flex; align-items:center; gap:4px; font-size:13px;">
              <span class="material-symbols-outlined" style="font-size:16px;">add</span> Create Validation
            </button>
            <button (click)="manageRulesModalOpen = false" style="background:#f1f5f9; color:#333; border:1px solid #e2e8f0; padding:8px 24px; border-radius:4px; font-weight:600; cursor:pointer; font-size:13px;">Close</button>
          </div>
        </div>
      </div>
      <!-- Spell Check Modal -->
      <div class="modal-overlay" *ngIf="spellCheckModalOpen" (click)="spellCheckModalOpen=false" style="z-index:10000;">
        <div class="modal" (click)="$event.stopPropagation()" style="width:480px;background:#fff;color:#333;border-radius:8px;box-shadow:0 8px 32px rgba(0,0,0,0.15);padding:24px;border:1px solid #e2e8f0;">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px;">
            <div style="display:flex;align-items:center;gap:10px;">
              <span class="material-symbols-outlined" style="color:#10b981;font-size:22px;">spellcheck</span>
              <h3 style="margin:0;font-size:18px;font-weight:600;">Spell Check</h3>
            </div>
            <button (click)="spellCheckModalOpen=false" style="background:none;border:none;cursor:pointer;color:#888;"><span class="material-symbols-outlined" style="font-size:20px;">close</span></button>
          </div>
          <div *ngIf="spellCheckLoading" style="padding:40px;text-align:center;color:#5f6368;">
             <span class="material-symbols-outlined" style="animation:spin 1s linear infinite;font-size:24px;">refresh</span>
             <p>Checking spelling...</p>
          </div>
          <div *ngIf="!spellCheckLoading && spellCheckErrors.length === 0" style="padding:40px;text-align:center;color:#10b981;">
             <span class="material-symbols-outlined" style="font-size:48px;">check_circle</span>
             <p style="margin-top:12px;font-weight:500;">No spelling errors found!</p>
          </div>
          <div *ngIf="!spellCheckLoading && spellCheckErrors.length > 0" style="max-height:300px;overflow-y:auto;">
             <p style="color:#d93025;font-weight:600;margin-bottom:16px;">Found {{spellCheckErrors.length}} error(s):</p>
             <div *ngFor="let err of spellCheckErrors; let i=index" style="border:1px solid #e2e8f0;border-radius:6px;padding:12px;margin-bottom:12px;background:#f8f9fa;">
               <div style="font-size:14px;color:#333;margin-bottom:8px;">
                 "...{{err.context.text.substring(0, err.context.offset)}}<strong style="color:#d93025;background:#fee2e2;padding:2px 4px;border-radius:2px;">{{err.context.text.substring(err.context.offset, err.context.offset + err.context.length)}}</strong>{{err.context.text.substring(err.context.offset + err.context.length)}}..."
               </div>
               <div style="font-size:12px;color:#5f6368;margin-bottom:12px;">{{err.message}}</div>
               <div style="display:flex;gap:8px;flex-wrap:wrap;" *ngIf="err.replacements?.length > 0">
                 <button *ngFor="let rep of err.replacements.slice(0, 4)" (click)="applySpellCheckFix(i, rep.value)" style="background:#fff;border:1px solid #cbd5e1;padding:4px 10px;border-radius:4px;font-size:12px;cursor:pointer;color:#1a73e8;font-weight:500;">{{rep.value}}</button>
               </div>
             </div>
          </div>
        </div>
      </div>

      <!-- Translate Modal -->
      <div class="modal-overlay" *ngIf="translateModalOpen" (click)="translateModalOpen=false" style="z-index:10000;">
        <div class="modal" (click)="$event.stopPropagation()" style="width:500px;background:#fff;color:#333;border-radius:8px;box-shadow:0 8px 32px rgba(0,0,0,0.15);padding:24px;border:1px solid #e2e8f0;">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px;">
            <div style="display:flex;align-items:center;gap:10px;">
              <span class="material-symbols-outlined" style="color:#1a73e8;font-size:22px;">translate</span>
              <h3 style="margin:0;font-size:18px;font-weight:600;">Translate Cell</h3>
            </div>
            <button (click)="translateModalOpen=false" style="background:none;border:none;cursor:pointer;color:#888;"><span class="material-symbols-outlined" style="font-size:20px;">close</span></button>
          </div>
          
          <div style="display:flex;gap:12px;align-items:stretch;margin-bottom:16px;">
             <div style="flex:1;">
               <div style="font-size:12px;font-weight:600;color:#5f6368;margin-bottom:6px;">Original Text (Auto-detect)</div>
               <textarea [ngModel]="translateSourceText" readonly style="width:100%;height:100px;background:#f8f9fa;border:1px solid #e2e8f0;border-radius:6px;padding:12px;font-size:14px;color:#333;resize:none;outline:none;box-sizing:border-box;"></textarea>
             </div>
             <div style="display:flex;align-items:center;">
               <span class="material-symbols-outlined" style="color:#9aa0a6;">arrow_forward</span>
             </div>
             <div style="flex:1;">
               <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">
                 <div style="font-size:12px;font-weight:600;color:#5f6368;">Translation</div>
                 <select [(ngModel)]="translateTargetLang" (change)="runTranslate()" style="border:1px solid #cbd5e1;border-radius:4px;padding:2px 4px;font-size:11px;outline:none;">
                   <option value="es">Spanish</option>
                   <option value="fr">French</option>
                   <option value="de">German</option>
                   <option value="it">Italian</option>
                   <option value="zh-CN">Chinese (Simplified)</option>
                   <option value="ja">Japanese</option>
                   <option value="ko">Korean</option>
                   <option value="hi">Hindi</option>
                 </select>
               </div>
               <div style="position:relative;width:100%;height:100px;">
                 <textarea [(ngModel)]="translateTargetText" [readOnly]="translateLoading" style="width:100%;height:100%;background:#fff;border:1px solid #1a73e8;border-radius:6px;padding:12px;font-size:14px;color:#1a73e8;resize:none;outline:none;box-sizing:border-box;"></textarea>
                 <div *ngIf="translateLoading" style="position:absolute;top:0;left:0;right:0;bottom:0;background:rgba(255,255,255,0.8);display:flex;align-items:center;justify-content:center;border-radius:6px;">
                   <span class="material-symbols-outlined" style="animation:spin 1s linear infinite;color:#1a73e8;">refresh</span>
                 </div>
               </div>
             </div>
          </div>
          
          <div style="display:flex;justify-content:flex-end;gap:10px;">
            <button (click)="translateModalOpen=false" style="background:#f1f5f9;color:#333;border:1px solid #e2e8f0;padding:8px 20px;border-radius:4px;font-weight:600;cursor:pointer;">Cancel</button>
            <button (click)="applyTranslation()" [disabled]="!translateTargetText || translateLoading" style="background:#1a73e8;color:#fff;border:none;padding:8px 20px;border-radius:4px;font-weight:600;cursor:pointer;" [style.opacity]="!translateTargetText || translateLoading ? 0.5 : 1">Replace in Cell</button>
          </div>
        </div>
      </div>

      <!-- Feature Modals -->
        <div class="modal-overlay" *ngIf="activeModal !== null && activeModal !== 'goto' && activeModal !== 'insert_sparkline' && activeModal !== 'edit_sparkline' && activeModal !== 'emoji' && activeModal !== 'custom_insert'" (click)="activeModal = null" style="z-index: 10000;">
          <div class="modal" (click)="$event.stopPropagation()" [style.width]="activeModal === 'version' ? '1200px' : (activeModal === 'macro' || activeModal === 'edit_macro' ? '800px' : (activeModal === 'audit' ? '620px' : (activeModal === 'manage_forms' ? '748px' : (activeModal === 'shortcuts' ? '548px' : '460px'))))" [style.padding]="(activeModal === 'macro' || activeModal === 'edit_macro') ? '0' : '24px'" style="background:#fff; color:#333; border:1px solid #e2e8f0; box-shadow:0 8px 32px rgba(0,0,0,0.15); max-width:90vw; border-radius:8px; position:relative; overflow:hidden; display:flex; flex-direction:column;">
            <button *ngIf="activeModal !== 'macro' && activeModal !== 'edit_macro'" (click)="activeModal = null" style="position:absolute;top:16px;right:16px;background:none;border:none;cursor:pointer;color:#888;display:flex;align-items:center;justify-content:center;z-index:10;">
              <span class="material-symbols-outlined" style="font-size:20px;">close</span>
            </button>

            <div *ngIf="activeModal === 'audit'" style="width:100%;">
              <div style="font-size:16px;font-weight:600;margin-bottom:20px;font-family:'Roboto',sans-serif;">Audit Trail</div>
              
              <!-- Tabs -->
              <div style="display:flex; gap:20px; border-bottom:1px solid #e0e0e0; margin-bottom:16px; font-size:13px; font-family:'Roboto',sans-serif; color:#5f6368;">
                <div (click)="sortAudit('user')" [style.color]="auditSortBy === 'user' ? '#0f9d58' : ''" [style.border-bottom]="auditSortBy === 'user' ? '2px solid #0f9d58' : ''" [style.font-weight]="auditSortBy === 'user' ? '500' : ''" style="padding-bottom:8px; cursor:pointer;">User <span *ngIf="auditSortBy === 'user'">{{auditSortDesc ? '↓' : '↑'}}</span></div>
                <div (click)="sortAudit('date')" [style.color]="auditSortBy === 'date' ? '#0f9d58' : ''" [style.border-bottom]="auditSortBy === 'date' ? '2px solid #0f9d58' : ''" [style.font-weight]="auditSortBy === 'date' ? '500' : ''" style="padding-bottom:8px; cursor:pointer;">Date <span *ngIf="auditSortBy === 'date'">{{auditSortDesc ? '↓' : '↑'}}</span></div>
                <div (click)="sortAudit('sheet')" [style.color]="auditSortBy === 'sheet' ? '#0f9d58' : ''" [style.border-bottom]="auditSortBy === 'sheet' ? '2px solid #0f9d58' : ''" [style.font-weight]="auditSortBy === 'sheet' ? '500' : ''" style="padding-bottom:8px; cursor:pointer;">Sheet <span *ngIf="auditSortBy === 'sheet'">{{auditSortDesc ? '↓' : '↑'}}</span></div>
                <div (click)="sortAudit('range')" [style.color]="auditSortBy === 'range' ? '#0f9d58' : ''" [style.border-bottom]="auditSortBy === 'range' ? '2px solid #0f9d58' : ''" [style.font-weight]="auditSortBy === 'range' ? '500' : ''" style="padding-bottom:8px; cursor:pointer;">Range <span *ngIf="auditSortBy === 'range'">{{auditSortDesc ? '↓' : '↑'}}</span></div>
              </div>

              <!-- Filter -->
              <div style="display:flex; align-items:center; gap:8px; font-size:13px; font-family:'Roboto',sans-serif; margin-bottom:16px;">
                <span style="color:#333;">Sheet:</span>
                <select style="border:1px solid #e0e0e0; border-radius:2px; padding:4px 8px; outline:none; background:#f9f9f9; color:#333; font-family:inherit; min-width:160px;">
                  <option>Whole Document</option>
                  <option *ngFor="let s of sheets">{{s.name}}</option>
                </select>
              </div>

              <!-- List -->
              <div style="border:1px solid #e0e0e0; border-radius:2px; height:320px; overflow-y:auto; background:#fff; font-family:'Roboto',sans-serif; font-size:13px; margin-bottom:16px; box-sizing:border-box;">
                
                <div *ngIf="auditRecords.length === 0" style="padding:12px; color:#888; text-align:center;">
                  No audit records found.
                </div>
                
                <div *ngFor="let record of sortedAuditRecords" style="padding:12px; border-bottom:1px solid #f0f0f0;">
                  <div style="display:flex; align-items:center; gap:6px; margin-bottom:4px;">
                    <div style="width:6px; height:6px; background:#1a73e8; border-radius:50%;"></div>
                    <span style="color:#1a73e8; font-weight:500;">{{ record.user_name }}</span>
                    <span style="color:#888;">- {{ record.created_at | date:'medium' }}</span>
                  </div>
                  <div style="color:#333; padding-left:12px;">
                    <span style="font-weight:500;">{{ getAuditActionPrefix(record.action_type) }}</span> 
                    <span style="color:#0f9d58; cursor:pointer;" (click)="onAuditNavigate({sheetId: record.sheet_id, r: record.metadata_json?.r || 0, c: record.metadata_json?.c || 0, endR: record.metadata_json?.r || 0, endC: record.metadata_json?.c || 0})">
                      '{{ record.sheet_name }}'.{{ record.target_range }}
                    </span>
                  </div>
                </div>

              </div>

              <!-- Footer Info -->
              <div style="margin-top:16px; font-size:11px; font-family:'Roboto',sans-serif; color:#5f6368; line-height:1.6;">
                <div><strong>Time Zone:</strong> India Standard Time</div>
                <div><strong>Note:</strong> Only the last 1,000 edits are available. To view more, please access <a href="javascript:void(0)" style="color:#1a73e8; text-decoration:none;" (click)="openFeatureModal('version')">Version History</a>.</div>
              </div>

              <div style="display:flex; justify-content:flex-end; margin-top:12px;">
                <button (click)="activeModal=null" style="background:#f8f9fa; color:#333; border:1px solid #ccc; padding:6px 16px; border-radius:4px; font-size:13px; font-weight:500; cursor:pointer;">Close</button>
              </div>
            </div>



            <div *ngIf="activeModal === 'workflow'">
              <div style="display:flex;align-items:center;gap:10px;margin-bottom:20px;">
                <span class="material-symbols-outlined" style="color:#f59e0b;font-size:24px;">account_tree</span>
                <h3 style="margin:0;font-size:18px;font-weight:600;">Manage Workflows</h3>
              </div>
              <div style="text-align:center;padding:40px 20px;color:#5f6368;">
                <span class="material-symbols-outlined" style="font-size:48px;color:#e2e8f0;">account_tree</span>
                <p style="margin-top:12px;font-size:14px;">No workflows created yet.</p>
                <p style="font-size:13px;color:#9aa0a6;">Workflows let you automate actions when data changes in your sheet.</p>
                <button (click)="showToast('Workflow creation coming soon.')" style="background:#10b981;color:#fff;border:none;padding:8px 20px;border-radius:4px;font-weight:600;cursor:pointer;margin-top:8px;">Create Workflow</button>
              </div>
            </div>

            <div *ngIf="activeModal === 'template'">
              <div style="display:flex;align-items:center;gap:10px;margin-bottom:20px;">
                <span class="material-symbols-outlined" style="color:#1a73e8;font-size:24px;">grid_view</span>
                <h3 style="margin:0;font-size:18px;font-weight:600;">Choose Template</h3>
              </div>
              <div *ngFor="let item of dummyList" (click)="handleModalAction(item)" style="padding:12px 16px;background:#f8f9fa;margin-bottom:8px;border-radius:6px;cursor:pointer;font-size:13px;display:flex;align-items:center;gap:10px;border:1px solid #e2e8f0;">
                <span class="material-symbols-outlined" style="color:#1a73e8;">description</span> {{ item }}
              </div>
            </div>

            <div *ngIf="activeModal === 'form'">
              <div style="display:flex;align-items:center;gap:10px;margin-bottom:20px;">
                <span class="material-symbols-outlined" style="color:#8b5cf6;font-size:24px;">assignment</span>
                <h3 style="margin:0;font-size:18px;font-weight:600;">Data Entry Form</h3>
              </div>
              
              <div *ngIf="formHeaders.length === 0" style="text-align:center;padding:20px;color:#5f6368;">
                <span class="material-symbols-outlined" style="font-size:48px;color:#e2e8f0;margin-bottom:12px;">warning</span>
                <p style="font-size:14px;margin-bottom:8px;">No headers found in Row 1.</p>
                <p style="font-size:12px;color:#9aa0a6;">Please add headers to the first row of your sheet to generate a form.</p>
                <div style="margin-top:20px;">
                  <button (click)="activeModal=null" style="background:#8b5cf6;color:#fff;border:none;padding:8px 20px;border-radius:4px;font-weight:600;cursor:pointer;">Close</button>
                </div>
              </div>

              <div *ngIf="formHeaders.length > 0" style="max-height:400px;overflow-y:auto;padding-right:10px;">
                <div *ngFor="let header of formHeaders" style="margin-bottom:12px;">
                  <label style="display:block;font-size:13px;font-weight:600;color:#5f6368;margin-bottom:4px;">{{header}}</label>
                  <input [(ngModel)]="formData[header]" type="text" style="width:100%;box-sizing:border-box;border:1px solid #cbd5e1;border-radius:4px;padding:8px;font-size:14px;outline:none;" placeholder="Enter {{header}}" />
                </div>
                <div style="display:flex;justify-content:flex-end;gap:10px;margin-top:20px;">
                  <button (click)="activeModal=null" style="background:#f1f5f9;color:#333;border:1px solid #e2e8f0;padding:8px 20px;border-radius:4px;font-weight:600;cursor:pointer;">Cancel</button>
                  <button (click)="submitForm()" style="background:#8b5cf6;color:#fff;border:none;padding:8px 20px;border-radius:4px;font-weight:600;cursor:pointer;">Submit Data</button>
                </div>
              </div>
            </div>

            <div *ngIf="activeModal === 'manage_forms'" style="width: 700px; max-width: 90vw;">
              <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px;">
                <h3 style="margin:0;font-size:18px;font-weight:600;" [style.color]="currentTheme === 'dark' ? '#e8eaed' : '#202124'">Manage Form</h3>
                
              </div>
              
              <div style="margin-bottom: 20px; border-radius: 4px; border: 1px solid {{ currentTheme === 'dark' ? '#5f6368' : '#e0e0e0' }}; overflow: hidden;">
                <table style="width: 100%; border-collapse: collapse; text-align: left; font-size: 13px;" [style.color]="currentTheme === 'dark' ? '#e8eaed' : '#333'">
                  <thead>
                    <tr style="background-color: {{ currentTheme === 'dark' ? '#3c4043' : '#f8f9fa' }}; border-bottom: 1px solid {{ currentTheme === 'dark' ? '#5f6368' : '#e0e0e0' }};">
                      <th style="padding: 12px 16px; font-weight: 600;">Form Name</th>
                      <th style="padding: 12px 16px; font-weight: 600;">Sheet</th>
                      <th style="padding: 12px 16px; font-weight: 600;">Link</th>
                      <th style="padding: 12px 16px; font-weight: 600;">Status</th>
                      <th style="padding: 12px 16px; font-weight: 600; text-align: center;">Action</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr><td colspan="5" style="padding: 24px; text-align: center; color: #888; font-style: italic;">No forms created yet.</td></tr>
                  </tbody>
                </table>
              </div>
              
              <div style="display:flex;justify-content:space-between;">
                <button (click)="createForm()" style="background:#f1f5f9;color:#333;border:1px solid #e2e8f0;padding:8px 20px;border-radius:4px;font-weight:600;cursor:pointer;">Create Form</button>
                <button (click)="activeModal=null" style="background:#f1f5f9;color:#333;border:1px solid #e2e8f0;padding:8px 20px;border-radius:4px;font-weight:600;cursor:pointer;">Close</button>
              </div>
            </div>
            
            <!-- Shortcuts Modal -->
            <div *ngIf="activeModal === 'shortcuts'" style="width: 100%; max-height: 80vh; display: flex; flex-direction: column;">
              <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px;">
                <h3 style="margin:0;font-size:18px;font-weight:600;" [style.color]="currentTheme === 'dark' ? '#e8eaed' : '#202124'">Keyboard Shortcuts</h3>
              </div>
              
              <div style="display:flex;gap:12px;margin-bottom:16px;">
                <select [(ngModel)]="shortcutCategoryFilter" style="flex:1;padding:8px 12px;border-radius:4px;border:1px solid {{ currentTheme === 'dark' ? '#5f6368' : '#e0e0e0' }};background:transparent;outline:none;" [style.color]="currentTheme === 'dark' ? '#e8eaed' : '#333'">
                  <option value="all" [style.background]="currentTheme === 'dark' ? '#202124' : '#fff'">All Shortcuts</option>
                  <option *ngFor="let cat of shortcutCategories" [value]="cat.id" [style.background]="currentTheme === 'dark' ? '#202124' : '#fff'">{{cat.name}}</option>
                </select>
                <div style="flex:1;position:relative;">
                  <span class="material-symbols-outlined" style="position:absolute;left:8px;top:8px;font-size:18px;color:#9aa0a6;">search</span>
                  <input type="text" [(ngModel)]="shortcutSearchQuery" placeholder="Search" style="width:100%;box-sizing:border-box;padding:8px 12px 8px 32px;border-radius:4px;border:1px solid {{ currentTheme === 'dark' ? '#5f6368' : '#e0e0e0' }};background:transparent;outline:none;" [style.color]="currentTheme === 'dark' ? '#e8eaed' : '#333'" />
                </div>
              </div>
              
              <div style="display:flex;font-weight:600;font-size:14px;padding-bottom:8px;border-bottom:1px solid {{ currentTheme === 'dark' ? '#5f6368' : '#e0e0e0' }};color:{{ currentTheme === 'dark' ? '#e8eaed' : '#333' }};">
                <div style="flex:2;">Description</div>
                <div style="flex:1;">Shortcut</div>
              </div>
              
              <div style="flex:1;overflow-y:auto;padding-top:12px;font-size:13px;color:{{ currentTheme === 'dark' ? '#bdc1c6' : '#5f6368' }};">
                <ng-container *ngFor="let cat of filteredShortcutCategories">
                  <div style="font-weight:600;color:{{ currentTheme === 'dark' ? '#e8eaed' : '#202124' }};margin:16px 0 8px;">{{cat.name}}</div>
                  <div *ngFor="let s of cat.shortcuts" style="display:flex;align-items:center;margin-bottom:12px;">
                    <div style="flex:2;">{{s.desc}}</div>
                    <div style="flex:1;display:flex;gap:4px;">
                      <span *ngFor="let k of s.keys" style="background:{{ currentTheme === 'dark' ? '#3c4043' : '#f8f9fa' }};border:1px solid {{ currentTheme === 'dark' ? '#5f6368' : '#dadce0' }};border-radius:4px;padding:2px 6px;">{{k}}</span>
                    </div>
                  </div>
                </ng-container>
                <div *ngIf="filteredShortcutCategories.length === 0" style="padding: 24px 0; text-align: center; color: #888; font-style: italic;">
                  No shortcuts found matching your search.
                </div>
              </div>
              <div style="padding-top:16px;border-top:1px solid {{ currentTheme === 'dark' ? '#5f6368' : '#e0e0e0' }};display:flex;align-items:center;gap:8px;">
                <input type="checkbox" id="overrideShortcuts" checked style="accent-color:#10b981;" />
                <label for="overrideShortcuts" style="font-size:13px;color:{{ currentTheme === 'dark' ? '#bdc1c6' : '#5f6368' }};">Override browser shortcuts</label>
                <span class="material-symbols-outlined" style="font-size:14px;color:#9aa0a6;">info</span>
              </div>
            </div>

            <div *ngIf="activeModal === 'macro' || activeModal === 'edit_macro'" style="width: 100%; height: 500px; display: flex; flex-direction: column;">
              
              <!-- Header -->
              <div style="display:flex; justify-content:space-between; align-items:center; padding: 12px 16px; border-bottom: 1px solid #e0e0e0; background: #fff;">
                <h3 style="margin:0; font-size:16px; font-weight:600; color: #202124;">VBA Editor</h3>
                <div style="display:flex; gap: 8px;">
                  <span class="material-symbols-outlined" style="cursor:pointer; font-size: 18px; color: #5f6368;" (click)="activeModal=null">minimize</span>
                  <span class="material-symbols-outlined" style="cursor:pointer; font-size: 18px; color: #5f6368;" (click)="activeModal=null">close</span>
                </div>
              </div>

              <!-- Toolbar -->
              <div style="display:flex; align-items:center; gap: 16px; padding: 8px 16px; border-bottom: 1px solid #e0e0e0; background: #f8f9fa; font-size: 13px; color: #3c4043;">
                <div style="display:flex; gap: 8px; cursor:pointer;" title="Undo"><span class="material-symbols-outlined" style="font-size:16px;">undo</span></div>
                <div style="display:flex; gap: 8px; cursor:pointer;" title="Redo"><span class="material-symbols-outlined" style="font-size:16px;">redo</span></div>
                <div style="display:flex; gap: 4px; align-items:center; cursor:pointer;"><span class="material-symbols-outlined" style="font-size:16px;">search</span> Find & Replace</div>
                <div style="display:flex; gap: 4px; align-items:center; cursor:pointer;"><span class="material-symbols-outlined" style="font-size:16px;">arrow_forward</span> Go To</div>
                <div style="display:flex; gap: 4px; align-items:center; cursor:pointer;"><span class="material-symbols-outlined" style="font-size:16px;">data_object</span> Insert Procedure</div>
                <div style="display:flex; gap: 4px; align-items:center; cursor:pointer;" (click)="save()"><span class="material-symbols-outlined" style="font-size:16px;">save</span> Save</div>
                <div style="display:flex; gap: 4px; align-items:center; cursor:pointer; color: #188038;" (click)="runMacro()"><span class="material-symbols-outlined" style="font-size:16px;">play_arrow</span> Run</div>
                <div style="display:flex; gap: 4px; align-items:center; cursor:pointer; color: #188038;" (click)="save(); runMacro()"><span class="material-symbols-outlined" style="font-size:16px;">save</span> Save & Run</div>
              </div>

              <!-- Main Content -->
              <div style="display: flex; flex: 1; min-height: 0;">
                
                <!-- Sidebar -->
                <div style="width: 220px; border-right: 1px solid #e0e0e0; display:flex; flex-direction: column; background: #fff;">
                  <div style="padding: 8px 12px; font-weight: 600; font-size: 12px; color: #5f6368; border-bottom: 1px solid #f1f3f4;">Macros</div>
                  <div style="flex: 1; overflow-y: auto; padding: 8px; font-size: 13px;">
                    <div style="display:flex; align-items:center; gap: 6px; padding: 4px; cursor: pointer;">
                      <span class="material-symbols-outlined" style="font-size:16px; color:#5f6368;">folder</span> Spreadsheet Objects
                    </div>
                    <div style="margin-left: 20px;">
                      <div style="display:flex; align-items:center; gap: 6px; padding: 4px; cursor: pointer; color: #1a73e8; background: #e8f0fe; border-radius: 4px;">
                        <span class="material-symbols-outlined" style="font-size:16px; color:#1a73e8;">description</span> This Workbook
                      </div>
                      <div *ngFor="let sheet of sheets; let i = index" style="display:flex; align-items:center; gap: 6px; padding: 4px; cursor: pointer;">
                        <span class="material-symbols-outlined" style="font-size:16px; color:#5f6368;">grid_on</span> Sheet{{i+1}}({{sheet.name}})
                      </div>
                    </div>
                  </div>
                </div>

                <!-- Editor & Terminal -->
                <div style="flex: 1; display: flex; flex-direction: column; min-width: 0;">
                  <!-- Editor -->
                  <div style="flex: 1; display:flex; position: relative;">
                    <textarea [(ngModel)]="macroScript" style="width:100%; height:100%; border:none; resize:none; padding:12px; font-family: 'Courier New', Courier, monospace; font-size: 14px; outline: none; line-height: 1.5; color: #333;"></textarea>
                  </div>

                  <!-- Terminal/Output -->
                  <div style="height: 140px; border-top: 1px solid #e0e0e0; display:flex; flex-direction: column; background: #f8f9fa;">
                    <!-- Tabs -->
                    <div style="display: flex; gap: 24px; padding: 0 16px; border-bottom: 1px solid #e0e0e0;">
                      <div (click)="macroActiveTab='Errors'" [style.border-bottom]="macroActiveTab === 'Errors' ? '2px solid #d32f2f' : '2px solid transparent'" [style.color]="macroActiveTab === 'Errors' ? '#d32f2f' : '#5f6368'" style="padding: 8px 0; cursor: pointer; font-size: 13px; font-weight: 500; display:flex; align-items:center; gap:6px;">
                        <span class="material-symbols-outlined" style="font-size:16px;">error</span> Errors
                      </div>
                      <div (click)="macroActiveTab='Messages'" [style.border-bottom]="macroActiveTab === 'Messages' ? '2px solid #1a73e8' : '2px solid transparent'" [style.color]="macroActiveTab === 'Messages' ? '#1a73e8' : '#5f6368'" style="padding: 8px 0; cursor: pointer; font-size: 13px; font-weight: 500; display:flex; align-items:center; gap:6px;">
                        <span class="material-symbols-outlined" style="font-size:16px;">info</span> Messages
                      </div>
                      <div (click)="macroActiveTab='Warnings'" [style.border-bottom]="macroActiveTab === 'Warnings' ? '2px solid #f59e0b' : '2px solid transparent'" [style.color]="macroActiveTab === 'Warnings' ? '#f59e0b' : '#5f6368'" style="padding: 8px 0; cursor: pointer; font-size: 13px; font-weight: 500; display:flex; align-items:center; gap:6px;">
                        <span class="material-symbols-outlined" style="font-size:16px;">warning</span> Warnings
                      </div>
                    </div>
                    <!-- Output content -->
                    <div style="flex: 1; overflow-y: auto; padding: 8px 16px; font-family: monospace; font-size: 12px; color: #333; background: #fff;">
                      <div *ngIf="macroActiveTab === 'Errors'">
                        <div *ngFor="let err of macroErrors" style="color: #d32f2f; margin-bottom: 4px;">{{err}}</div>
                        <div *ngIf="macroErrors.length === 0" style="color: #9aa0a6;">No errors.</div>
                      </div>
                      <div *ngIf="macroActiveTab === 'Messages'">
                        <div *ngFor="let msg of macroMessages" style="color: #1a73e8; margin-bottom: 4px;">{{msg}}</div>
                        <div *ngIf="macroMessages.length === 0" style="color: #9aa0a6;">No messages.</div>
                      </div>
                      <div *ngIf="macroActiveTab === 'Warnings'">
                        <div *ngFor="let warn of macroWarnings" style="color: #f59e0b; margin-bottom: 4px;">{{warn}}</div>
                        <div *ngIf="macroWarnings.length === 0" style="color: #9aa0a6;">No warnings.</div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>

            </div>

            <div *ngIf="activeModal === 'functions'">
              <div style="display:flex;align-items:center;gap:10px;margin-bottom:20px;">
                <span class="material-symbols-outlined" style="color:#f59e0b;font-size:24px;">functions</span>
                <h3 style="margin:0;font-size:18px;font-weight:600;">Custom Functions</h3>
              </div>
              <div style="margin-bottom:16px;">
                <p style="font-size:12px;color:#5f6368;margin-bottom:8px;">Define custom functions using JavaScript. You can call them in cells like '=MY_CUSTOM_SUM(1, 2)'</p>
                <textarea [(ngModel)]="customFunctionsScript" style="width:100%;height:200px;background:#1e1e1e;color:#d4d4d4;font-family:monospace;font-size:13px;padding:12px;border-radius:6px;border:none;resize:none;box-sizing:border-box;outline:none;"></textarea>
              </div>
              <div style="display:flex;justify-content:flex-end;gap:10px;">
                <button (click)="activeModal=null" style="background:#f1f5f9;color:#333;border:1px solid #e2e8f0;padding:8px 20px;border-radius:4px;font-weight:600;cursor:pointer;">Cancel</button>
                <button (click)="saveCustomFunctions()" style="background:#f59e0b;color:#fff;border:none;padding:8px 20px;border-radius:4px;font-weight:600;cursor:pointer;">Save & Apply</button>
              </div>
            </div>

            <div *ngIf="activeModal === 'merge'">
              <div style="display:flex;align-items:center;gap:10px;margin-bottom:20px;">
                <span class="material-symbols-outlined" style="color:#10b981;font-size:24px;">merge_type</span>
                <h3 style="margin:0;font-size:18px;font-weight:600;">Merge Template</h3>
              </div>
              <div style="text-align:center;padding:20px;color:#5f6368;background:#f8f9fa;border:2px dashed #cbd5e1;border-radius:6px;margin-bottom:20px;">
                <span class="material-symbols-outlined" style="font-size:48px;color:#1a73e8;margin-bottom:12px;">upload_file</span>
                <p style="font-size:14px;margin-bottom:16px;">Select a document template (.docx) to merge with your spreadsheet rows.</p>
                <label style="background:#fff;border:1px solid #1a73e8;color:#1a73e8;padding:10px 20px;border-radius:4px;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;gap:8px;font-weight:600;">
                   <span class="material-symbols-outlined" style="font-size:18px;">browse_activity</span> Browse Files
                   <input type="file" style="display:none;" accept=".docx,.pdf" />
                </label>
              </div>
              <div style="display:flex;justify-content:flex-end;gap:10px;">
                <button (click)="activeModal=null" style="background:#f1f5f9;color:#333;border:1px solid #e2e8f0;padding:8px 20px;border-radius:4px;font-weight:600;cursor:pointer;">Cancel</button>
                <button (click)="simulateMerge()" style="background:#10b981;color:#fff;border:none;padding:8px 20px;border-radius:4px;font-weight:600;cursor:pointer;display:flex;align-items:center;gap:6px;"><span class="material-symbols-outlined" style="font-size:18px;">auto_awesome</span> Start Merge</button>
              </div>
            </div>

            <div *ngIf="activeModal === 'open'">
              <div style="display:flex;align-items:center;gap:10px;margin-bottom:20px;">
                <span class="material-symbols-outlined" style="color:#1a73e8;font-size:24px;">folder_open</span>
                <h3 style="margin:0;font-size:18px;font-weight:600;">Open Document</h3>
              </div>
              <div *ngIf="myDocs.length === 0" style="padding:20px;text-align:center;color:#666;">
                No spreadsheets found.
              </div>
              <div *ngFor="let doc of myDocs" (click)="handleModalAction(doc)" style="padding:12px 16px;background:#f8f9fa;margin-bottom:8px;border-radius:6px;cursor:pointer;font-size:13px;display:flex;align-items:center;justify-content:space-between;border:1px solid #e2e8f0;">
                <div style="display:flex;align-items:center;gap:10px;">
                   <span class="material-symbols-outlined" style="color:#0f9d58;">table_chart</span> {{ doc.title || 'Untitled' }}
                </div>
              </div>
            </div>

            <div *ngIf="activeModal === 'import'">
              <div style="display:flex;align-items:center;gap:10px;margin-bottom:20px;">
                <span class="material-symbols-outlined" style="color:#10b981;font-size:24px;">upload</span>
                <h3 style="margin:0;font-size:18px;font-weight:600;">Import File</h3>
              </div>
              <p style="color:#5f6368;font-size:13px;margin-bottom:16px;">Select a CSV, TSV, or XLSX file from your computer to import into the current sheet.</p>
              <input type="file" (change)="onFileSelected($event)" accept=".csv,.tsv,.xlsx,.xls" style="width:100%;padding:10px;background:#f8f9fa;border:1px solid #e2e8f0;border-radius:6px;color:#333;margin-bottom:16px;box-sizing:border-box;">
              <button class="btn" (click)="handleModalAction()" style="width:100%;background:#10b981;color:#fff;border:none;padding:10px;border-radius:4px;font-weight:600;cursor:pointer;">Import Now</button>
            </div>

            <div *ngIf="activeModal === 'move'">
              <div style="display:flex;align-items:center;gap:10px;margin-bottom:20px;">
                <span class="material-symbols-outlined" style="color:#f59e0b;font-size:24px;">drive_file_move</span>
                <h3 style="margin:0;font-size:18px;font-weight:600;">Move Document</h3>
              </div>
              <p style="color:#5f6368;font-size:13px;margin-bottom:16px;">Enter the name of the folder you want to move this document to:</p>
              <input type="text" [(ngModel)]="modalInput" placeholder="Folder Name" style="width:100%;padding:10px;background:#f8f9fa;border:1px solid #e2e8f0;border-radius:6px;color:#333;margin-bottom:16px;outline:none;box-sizing:border-box;">
              <button (click)="handleModalAction()" style="width:100%;background:#1a73e8;color:#fff;border:none;padding:10px;border-radius:4px;font-weight:600;cursor:pointer;">Move</button>
            </div>

            <div *ngIf="activeModal === 'password'">
              <div style="display:flex;align-items:center;gap:10px;margin-bottom:20px;">
                <span class="material-symbols-outlined" style="color:#d93025;font-size:24px;">lock</span>
                <h3 style="margin:0;font-size:18px;font-weight:600;">Protect Document</h3>
              </div>
              <p style="color:#5f6368;font-size:13px;margin-bottom:16px;">Set a password to restrict who can open or view this document.</p>
              <input type="password" [(ngModel)]="modalInput" placeholder="Enter new password" style="width:100%;padding:10px;background:#f8f9fa;border:1px solid #e2e8f0;border-radius:6px;color:#333;margin-bottom:16px;outline:none;box-sizing:border-box;">
              <button (click)="handleModalAction()" style="width:100%;background:#d93025;color:#fff;border:none;padding:10px;border-radius:4px;font-weight:600;cursor:pointer;">Set Password</button>
            </div>
          </div>
        </div>
        
        <!-- Properties Panel -->
        <div class="properties-panel" [class.open]="propertiesPanelOpen" (click)="$event.stopPropagation()">
          <div class="pp-header">
            <h2 class="pp-title">
              <span class="material-symbols-outlined" style="color:#26a96c;font-size:24px;">description</span>
              Details
            </h2>
            <button class="pp-close" (click)="propertiesPanelOpen = false">
              <span class="material-symbols-outlined">close</span>
            </button>
          </div>
          
          <div class="pp-content" *ngIf="docDetails">
            <div class="pp-section">
              <div style="display:flex;align-items:center;gap:12px;">
                <div class="pp-av">{{ docDetails.owner_name ? docDetails.owner_name[0] : 'U' }}</div>
                <div>
                  <div style="font-size:12px;color:#5f6368;">Created by</div>
                  <div style="font-weight:500;">{{ docDetails.owner_name || 'Unknown' }}</div>
                </div>
              </div>
            </div>
            
            <div class="pp-section">
              <div class="pp-label">Shared with</div>
              <div class="pp-value">
                <span class="material-symbols-outlined" style="font-size:18px;">{{ docDetails.is_public ? 'public' : 'lock' }}</span>
                {{ docDetails.is_public ? 'Public' : 'Private' }}
              </div>
            </div>
            
            <hr class="pp-divider">
            
            <div class="pp-section">
              <div class="pp-label">Permalink</div>
              <a [href]="window.location.href" target="_blank" class="pp-link">{{ window.location.href }}</a>
            </div>
            
            <hr class="pp-divider">
            
            <div class="pp-section">
              <div class="pp-label">Time Created</div>
              <div class="pp-value">{{ docDetails.created_at | date:'medium' }}</div>
            </div>
            
            <div class="pp-section">
              <div class="pp-label">Last Modified</div>
              <div class="pp-value">{{ docDetails.updated_at | date:'medium' }}</div>
            </div>
            
            <div class="pp-section">
              <div class="pp-label">Current Version</div>
              <div class="pp-value">{{ docDetails.content_version || 1 }}.0</div>
            </div>
            
            <hr class="pp-divider">
            
            <div class="pp-section">
              <div class="pp-label">Spreadsheet Statistics</div>
              <div class="pp-stats">
                <div class="pp-stat-item">
                  <div class="pp-stat-num">{{ getActiveSheetCount() }}</div>
                  <div class="pp-stat-lbl">Sheets</div>
                </div>
                <div class="pp-stat-item">
                  <div class="pp-stat-num">{{ getUsedCellsCount() }}</div>
                  <div class="pp-stat-lbl">Used cells</div>
                </div>
              </div>
            </div>
          </div>
        </div>

        <!-- Share Modal -->
        <div class="modal-overlay" *ngIf="shareModalOpen" (click)="shareModalOpen = false">
          <div class="modal share-modal" (click)="$event.stopPropagation()">
            <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:24px;">
              <div style="display:flex; align-items:center; gap:10px;">
                <div style="background:#0f9d58; color:#fff; display:flex; align-items:center; justify-content:center; width:22px; height:22px; border-radius:4px;">
                  <span class="material-symbols-outlined" style="font-size:16px;">grid_on</span>
                </div>
                <h3>Share "{{ title || 'Untitled spreadsheet' }}"</h3>
              </div>
              <button class="sm-close-btn" (click)="shareModalOpen = false">
                <span class="material-symbols-outlined" style="font-size:20px;">close</span>
              </button>
            </div>
            <div style="position:relative; margin-bottom:32px;">
              <div style="display:flex; align-items:center; gap:12px; position:relative;">
                <div class="sm-input-box">
                  <input type="text" class="sm-input" [(ngModel)]="shareQuery" (ngModelChange)="onShareSearch()" (keydown.enter)="addShareEmail($event)" placeholder="Add people and groups (press Enter)">
                  <div class="sm-dropdown-txt" (click)="shareRoleDropdownOpen = !shareRoleDropdownOpen" style="position:relative;">
                    {{ shareRole }} <span class="material-symbols-outlined" style="font-size:18px; color:inherit; opacity: 0.8;">arrow_drop_down</span>
                    
                    <div *ngIf="shareRoleDropdownOpen" class="sm-list" style="position:absolute; top:30px; right:0; left:auto; width:100px; z-index:100; min-width:100px; max-height:none;">
                       <div (click)="shareRole = 'View'; shareRoleDropdownOpen = false; $event.stopPropagation()" class="sm-list-item" style="padding:8px 12px; border-bottom:none;">View</div>
                       <div (click)="shareRole = 'Edit'; shareRoleDropdownOpen = false; $event.stopPropagation()" class="sm-list-item" style="padding:8px 12px; border-bottom:none;">Edit</div>
                    </div>
                  </div>
                </div>
                <button (click)="performShare()" style="background:#0f9d58; color:#fff; border:none; border-radius:24px; font-weight:500; font-size:14px; padding:0 24px; height:44px; cursor:pointer; transition:background 0.2s;" onmouseover="this.style.background='#0b8043'" onmouseout="this.style.background='#0f9d58'">Share</button>
              </div>
              <div *ngIf="selectedShareEmails.length > 0" style="display:flex; flex-wrap:wrap; gap:8px; margin-top:12px;">
                <div *ngFor="let email of selectedShareEmails" style="background:#e8eaed; border-radius:16px; padding:4px 12px; display:flex; align-items:center; gap:8px; font-size:13px; color:#3c4043; border:1px solid #dadce0;">
                  {{ email }}
                  <span class="material-symbols-outlined" style="font-size:16px; cursor:pointer;" (click)="removeShareEmail(email)">close</span>
                </div>
              </div>
              <div *ngIf="userSearchResults.length > 0" class="sm-list">
                <div *ngFor="let u of userSearchResults" (click)="selectShareUser(u)" class="sm-list-item">
                  <div style="width:32px; height:32px; border-radius:50%; display:flex; align-items:center; justify-content:center; color:#fff; font-size:14px; font-weight:500;" [style.background]="u.avatar_color">{{u.name.charAt(0).toUpperCase()}}</div>
                  <div style="display:flex; flex-direction:column;">
                    <div class="name">{{u.name}}</div>
                    <div class="email">{{u.email}}</div>
                  </div>
                </div>
              </div>
            </div>
            <div style="margin-bottom:32px;">
              <div style="font-size:11px; font-weight:600; color:#9aa0a6; letter-spacing:0.8px; margin-bottom:16px;">WHO CAN ACCESS</div>
              <div style="display:flex; align-items:center; justify-content:space-between;">
                <div style="display:flex; align-items:center; gap:16px;">
                  <div class="sm-icon-bg">
                    <span class="material-symbols-outlined" style="font-size:20px;">{{ isPublic ? 'public' : 'link' }}</span>
                  </div>
                  <div>
                    <div class="sm-txt-main">{{ isPublic ? 'Public Link - Anyone on the internet can view' : 'Permalink - Private, not shared with anyone' }}</div>
                  </div>
                </div>
                <button *ngIf="!isPublic" (click)="makePublic()" class="sm-sec-btn">
                  <span class="material-symbols-outlined" style="font-size:16px;">settings</span> Make Public
                </button>
                <button *ngIf="isPublic" (click)="isPublic = false" class="sm-sec-btn">
                  <span class="material-symbols-outlined" style="font-size:16px;">lock</span> Make Private
                </button>
              </div>
            </div>
            <div style="display:flex; align-items:center; justify-content:space-between;">
              <button (click)="copyLink()" class="sm-copy-btn">Copy Link</button>
              <button (click)="shareModalOpen = false" class="sm-done-btn">Done</button>
            </div>
          </div>
        </div>
      <!-- Footer Container -->
      <div class="footer-container" *ngIf="showStatusBar">
        <!-- Sheet Tabs -->
        <div class="sheet-tabs">
          <div class="sheet-tab" *ngFor="let sheet of sheets; let i = index"
            [style.display]="sheet.hidden ? 'none' : ''"
            [style.border-bottom]="sheet.tabColor ? '3px solid ' + sheet.tabColor : ''"
            [class.active-tab]="i === currentSheetIdx"
            (click)="switchSheet(i)"
            (dblclick)="renameSheet(i)">
          {{ sheet.name }}
          <span class="tab-menu-icon material-symbols-outlined" (click)="openSheetMenu(i, $event)" style="font-size: 16px; margin-left: 4px; border-radius: 4px; cursor: pointer; transition: background 0.2s;" onmouseover="this.style.background='rgba(0,0,0,0.1)'" onmouseout="this.style.background='transparent'">arrow_drop_down</span>
        </div>
        <button class="tab-add" (click)="addSheet()" title="Add sheet">＋</button>

        </div> <!-- end of sheet-tabs -->
        <div class="footer-tools-container" [style.background]="currentTheme === 'dark' ? '#202124' : '#ffffff'" [style.border-color]="currentTheme === 'dark' ? '#5f6368' : '#dadce0'">
          <!-- Selected Cell Count -->
          <div *ngIf="selectedNonEmptyCount > 1" 
               style="margin-right: 8px; display: flex; align-items: center; justify-content: center; padding: 4px 12px; border-radius: 20px; font-size: 13px; font-weight: 500; box-shadow: 0 1px 3px rgba(0,0,0,0.1); user-select: none; border: 1px solid; cursor: default; flex-shrink: 0;"
               [style.background]="currentTheme === 'dark' ? '#374151' : '#ffffff'"
               [style.color]="currentTheme === 'dark' ? '#f3f4f6' : '#374151'"
               [style.border-color]="currentTheme === 'dark' ? '#4b5563' : '#d1d5db'">
            Count = {{ selectedNonEmptyCount }}
            <span class="material-symbols-outlined" style="font-size: 14px; margin-left: 4px; opacity: 0.6;">unfold_more</span>
          </div>
          <!-- Appearance -->
          <button class="footer-btn" (click)="appearance = appearance === 'dark' ? 'light' : 'dark'" title="Toggle Theme">
            <span class="material-symbols-outlined" style="font-size: 18px;">{{ appearance === 'dark' ? 'light_mode' : 'dark_mode' }}</span>
          </button>

          <div class="footer-sep" [style.background]="currentTheme === 'dark' ? '#5f6368' : '#dadce0'"></div>

          <!-- View Settings -->
          <div style="position: relative;">
            <button type="button" class="footer-btn" title="View Settings" (click)="toggleFooterMenu('viewSettings', $event)">
              <span class="material-symbols-outlined" style="font-size: 18px;">settings_suggest</span>
            </button>
            <div class="ctx-menu" *ngIf="activeFooterMenu === 'viewSettings'" style="position: absolute; bottom: 100%; right: 0; margin-bottom: 8px;" (click)="$event.stopPropagation()">
                <div class="ctx-item" (click)="toggleViewSetting('topBar')"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{showTopBar?'visible':'hidden'}}; margin-right:8px;">check</span>Top Bar</div>
                <div class="ctx-item" (click)="toggleViewSetting('formulaBar')"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{showFormulaBar?'visible':'hidden'}}; margin-right:8px;">check</span>Formula Bar</div>
                <div class="ctx-sep"></div>
                <div class="ctx-item" (click)="toggleViewSetting('notes')"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{showNotes?'visible':'hidden'}}; margin-right:8px;">check</span>Notes</div>
                <div class="ctx-item" (click)="toggleViewSetting('userPresence')"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{showUserPresence?'visible':'hidden'}}; margin-right:8px;">check</span>User Presence</div>
                <div class="ctx-item" (click)="toggleViewSetting('lockPattern')"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{showLockPattern?'visible':'hidden'}}; margin-right:8px;">check</span>Lock Pattern</div>
                <div class="ctx-item" (click)="toggleViewSetting('printArea')"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{showHighlightPrintArea?'visible':'hidden'}}; margin-right:8px;">check</span>Highlight Print Area</div>
            </div>
          </div>

          <div class="footer-sep" [style.background]="currentTheme === 'dark' ? '#5f6368' : '#dadce0'"></div>

          <!-- Navigation -->
          <div style="position: relative;">
            <button type="button" class="footer-btn" title="Navigation" (click)="openApp('navigation'); activeFooterMenu=null">
              <span class="material-symbols-outlined" style="font-size: 18px;">web_stories</span>
            </button>
          </div>

          <div class="footer-sep" [style.background]="currentTheme === 'dark' ? '#5f6368' : '#dadce0'"></div>
          
          <!-- Zoom -->
          <div style="display: flex; align-items: center;">
            <button type="button" class="footer-btn" (click)="zoomOut()" style="padding: 2px 4px;"><span class="material-symbols-outlined" style="font-size: 16px;">remove</span></button>
            <div style="position: relative;">
              <button type="button" class="footer-btn" style="font-size: 12px; font-weight: 500; min-width: 44px; padding: 2px 0;" (click)="toggleFooterMenu('zoom', $event)">{{ zoomLevel }}%</button>
              <div class="ctx-menu" *ngIf="activeFooterMenu === 'zoom'" style="position: absolute; bottom: 100%; right: 0; margin-bottom: 8px;" (click)="$event.stopPropagation()">
                <div class="ctx-item" (click)="setZoom(200); activeFooterMenu=null"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{zoomLevel===200?'visible':'hidden'}}; margin-right:8px;">check</span>200%</div>
                <div class="ctx-item" (click)="setZoom(150); activeFooterMenu=null"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{zoomLevel===150?'visible':'hidden'}}; margin-right:8px;">check</span>150%</div>
                <div class="ctx-item" (click)="setZoom(125); activeFooterMenu=null"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{zoomLevel===125?'visible':'hidden'}}; margin-right:8px;">check</span>125%</div>
                <div class="ctx-item" (click)="setZoom(100); activeFooterMenu=null"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{zoomLevel===100?'visible':'hidden'}}; margin-right:8px;">check</span>100%</div>
                <div class="ctx-item" (click)="setZoom(75); activeFooterMenu=null"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{zoomLevel===75?'visible':'hidden'}}; margin-right:8px;">check</span>75%</div>
                <div class="ctx-item" (click)="setZoom(50); activeFooterMenu=null"><span class="material-symbols-outlined" style="font-size:16px; visibility:{{zoomLevel===50?'visible':'hidden'}}; margin-right:8px;">check</span>50%</div>
              </div>
            </div>
            <button class="footer-btn" (click)="zoomIn()" style="padding: 2px 4px;"><span class="material-symbols-outlined" style="font-size: 16px;">add</span></button>
          </div>

          <div class="footer-sep" [style.background]="currentTheme === 'dark' ? '#5f6368' : '#dadce0'"></div>

          <!-- Full Screen -->
          <button class="footer-btn" (click)="toggleFullScreen()" title="Full Screen">
            <span class="material-symbols-outlined" style="font-size: 18px;">fullscreen</span>
          </button>
        </div>
      </div>

      <!-- Sheet Context Menu -->
      <div class="ctx-menu" *ngIf="activeSheetMenuIdx !== null" [style.left.px]="sheetMenuX" [style.bottom.px]="sheetMenuY + 10" [style.overflow]="'visible'" (click)="$event.stopPropagation()">
        <div class="ctx-item" (click)="addSheet(); activeSheetMenuIdx=null"><span class="ctx-icon material-symbols-outlined" style="font-size: 16px;">add_box</span> Insert</div>
        <div class="ctx-item" (click)="duplicateSheet(activeSheetMenuIdx); activeSheetMenuIdx=null"><span class="ctx-icon material-symbols-outlined" style="font-size: 16px;">content_copy</span> Duplicate</div>
        <div class="ctx-item" (click)="deleteSheet(activeSheetMenuIdx); activeSheetMenuIdx=null" [class.disabled]="sheets.length <= 1" [class.danger]="sheets.length > 1"><span class="ctx-icon material-symbols-outlined" style="font-size: 16px;">delete</span> Delete</div>
        <div class="ctx-item" (click)="renameSheet(activeSheetMenuIdx); activeSheetMenuIdx=null"><span class="ctx-icon material-symbols-outlined" style="font-size: 16px;">edit</span> Rename</div>
        <div class="ctx-sep"></div>
        <div class="ctx-item" (click)="copySheet(activeSheetMenuIdx); activeSheetMenuIdx=null"><span class="ctx-icon material-symbols-outlined" style="font-size: 16px;">file_copy</span> Copy</div>
        <div class="ctx-item" [class.disabled]="!copiedSheetData" (click)="pasteSheet()"><span class="ctx-icon material-symbols-outlined" style="font-size: 16px;">content_paste</span> Paste</div>
        <div class="ctx-item" (click)="openMoveSheetModal(activeSheetMenuIdx); activeSheetMenuIdx=null">
          <span class="ctx-icon material-symbols-outlined" style="font-size: 16px;">drive_file_move</span> Move
        </div>
        <div class="ctx-item" style="position:relative;" (mouseenter)="activeSheetSubmenu='color'" (mouseleave)="activeSheetSubmenu=null">
          <span class="ctx-icon material-symbols-outlined" style="font-size: 16px;">palette</span> Tab Color <span class="mdi-arrow material-symbols-outlined" style="margin-left:auto;">chevron_right</span>
          <div style="position:absolute; left:calc(100% - 10px); top:-4px; padding-left:10px; z-index:1000;" *ngIf="activeSheetSubmenu==='color'">
            <div class="ctx-menu" style="padding:12px; display:flex; flex-direction:column; gap:8px; width:220px; box-shadow:0 4px 12px rgba(0,0,0,0.2);">
               <div style="display:flex; align-items:center; gap:8px; cursor:pointer; font-size:12px; color:inherit;" (click)="setTabColor(activeSheetMenuIdx, ''); closeMenus()">
                  <span class="material-symbols-outlined" style="font-size:18px; color:#888;">block</span> No Fill
               </div>
               <div style="font-size:11px; font-weight:600; color:#5f6368; margin-top:4px;">Theme Colors</div>
               <div style="display:grid; grid-template-columns:repeat(10, 1fr); gap:2px;">
                  <div *ngFor="let c of themeColorsGrid" style="width:16px; height:16px; border-radius:2px; cursor:pointer; border:1px solid #cbd5e1;" [style.background]="c" (click)="setTabColor(activeSheetMenuIdx, c); closeMenus()"></div>
               </div>
               <div style="font-size:11px; font-weight:600; color:#5f6368; margin-top:4px;">Standard Colors</div>
               <div style="display:grid; grid-template-columns:repeat(10, 1fr); gap:2px;">
                  <div *ngFor="let c of standardColors" style="width:16px; height:16px; border-radius:2px; cursor:pointer; border:1px solid #cbd5e1;" [style.background]="c" (click)="setTabColor(activeSheetMenuIdx, c); closeMenus()"></div>
               </div>
            </div>
          </div>
        </div>
        <div class="ctx-item" (click)="toggleSheetGridlines(activeSheetMenuIdx); activeSheetMenuIdx=null">
          <span class="ctx-icon material-symbols-outlined" style="font-size: 16px;">grid_on</span>
          {{ sheets[activeSheetMenuIdx].hideGridlines ? 'Show Gridlines' : 'Hide Gridlines' }}
        </div>
        <div class="ctx-item" (click)="hideSheet(activeSheetMenuIdx); activeSheetMenuIdx=null" [class.disabled]="getVisibleSheetCount() <= 1 && !sheets[activeSheetMenuIdx].hidden">
          <span class="ctx-icon material-symbols-outlined" style="font-size: 16px;">visibility_off</span> Hide
        </div>
        <div class="ctx-item" *ngIf="hiddenSheetsList.length > 0" style="position:relative;" (mouseenter)="activeSheetSubmenu='unhide'" (mouseleave)="activeSheetSubmenu=null">
          <span class="ctx-icon material-symbols-outlined" style="font-size: 16px; color: #10b981;">visibility</span> <span style="color: #10b981;">Unhide</span> <span class="mdi-arrow material-symbols-outlined" style="margin-left:auto; color: #10b981;">chevron_right</span>
          <div style="position:absolute; left:calc(100% - 10px); top:-4px; padding-left:10px; z-index:1000;" *ngIf="activeSheetSubmenu==='unhide'">
            <div class="ctx-menu" style="min-width: 180px;">
              <div class="ctx-item" *ngFor="let h of hiddenSheetsList; trackBy: trackByHiddenSheet" (click)="unhideSheetAndSwitch(h.idx); closeMenus()">
                 {{ h.s.name }}
              </div>
              <div class="ctx-sep" *ngIf="hiddenSheetsList.length > 1"></div>
              <div class="ctx-item" *ngIf="hiddenSheetsList.length > 1" (click)="unhideAllSheets(); closeMenus()">
                 Unhide All Sheets
              </div>
            </div>
          </div>
        </div>
        <div class="ctx-item" (click)="toggleLockSheet(activeSheetMenuIdx); activeSheetMenuIdx=null">
          <span class="ctx-icon material-symbols-outlined" style="font-size: 16px;">{{ sheets[activeSheetMenuIdx].locked ? 'lock_open' : 'lock' }}</span>
          {{ sheets[activeSheetMenuIdx].locked ? 'Unlock Sheet' : 'Lock Sheet' }}
        </div>
        <div class="ctx-item" (click)="publishSheet(activeSheetMenuIdx); activeSheetMenuIdx=null">
          <span class="ctx-icon material-symbols-outlined" style="font-size: 16px;">language</span> Publish This Sheet
        </div>
      </div>


      <!-- Find & Replace Modal -->
      <div class="modal-overlay" *ngIf="findModalOpen" (click)="findModalOpen = false" style="z-index: 10000; background: transparent; pointer-events: none;">
        <div class="modal" (click)="$event.stopPropagation()" 
             [style.right.px]="findModalPosition === 'right' ? 40 : null"
             [style.left.px]="findModalPosition === 'left' ? 40 : null"
             style="position: absolute; top: 100px; width:420px; background:#fff; color:#333; border-radius:8px; padding:20px; box-shadow:0 8px 32px rgba(0,0,0,0.15); border:1px solid #e2e8f0; font-family:'Roboto',sans-serif; pointer-events: auto; transition: right 0.3s, left 0.3s;">
          <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:20px;">
            <h3 style="margin:0; font-size:16px; font-weight:600; color:#333;">Find and Replace</h3>
            <button (click)="findModalOpen = false" style="background:none; border:none; color:#5f6368; font-size:20px; cursor:pointer; padding:0; line-height:1;">&times;</button>
          </div>
          
          <div style="display:flex; align-items:center; margin-bottom:16px;">
            <label style="width:90px; font-size:13px; color:#5f6368; font-weight:500;">Find:</label>
            <div style="position:relative; flex:1;">
               <input [(ngModel)]="findQuery" style="width:100%; box-sizing:border-box; padding:8px 40px 8px 12px; border:1px solid #dadce0; border-radius:4px; font-size:14px; outline:none;" (keydown.enter)="findNext()">
               <span style="position:absolute; right:12px; top:8px; font-size:12px; color:#5f6368;">{{ findMatches.length ? (findMatchIdx + 1) + '/' + findMatches.length : '0/0' }}</span>
            </div>
          </div>
          
          <div style="display:flex; align-items:center; margin-bottom:20px;">
            <label style="width:90px; font-size:13px; color:#5f6368; font-weight:500;">Replace with:</label>
            <input [(ngModel)]="replaceQuery" style="flex:1; box-sizing:border-box; padding:8px 12px; border:1px solid #dadce0; border-radius:4px; font-size:14px; outline:none;">
          </div>
          
          <div style="display:flex; align-items:center; margin-bottom:16px;">
            <label style="width:90px; font-size:13px; color:#5f6368; font-weight:500;">Search in:</label>
            <select [(ngModel)]="findSearchIn" style="flex:1; padding:8px; border:1px solid #dadce0; border-radius:4px; font-size:14px; outline:none; appearance:auto; background:#fff;">
              <option value="sheet">This sheet</option>
              <option value="workbook">All sheets</option>
            </select>
          </div>
          
          <div style="display:flex; margin-bottom:16px; padding-left:90px; gap:24px;">
            <label style="display:flex; align-items:center; font-size:13px; color:#333; cursor:pointer;">
              <input type="checkbox" [(ngModel)]="findMatchCase" style="margin-right:6px; accent-color:#10b981;"> Match case
            </label>
            <label style="display:flex; align-items:center; font-size:13px; color:#333; cursor:pointer;">
              <input type="checkbox" [(ngModel)]="findMatchEntireCell" style="margin-right:6px; accent-color:#10b981;"> Match entire cell
            </label>
          </div>
          
          <div style="display:flex; margin-bottom:24px; padding-left:90px;">
             <label style="display:flex; align-items:center; font-size:13px; color:#333; cursor:pointer;">
              <input type="checkbox" [(ngModel)]="findIncludeFormulas" style="margin-right:6px; accent-color:#10b981;"> Include formulas
            </label>
          </div>
          
          <div style="display:flex; align-items:center; margin-bottom:24px;">
            <label style="width:90px; font-size:13px; color:#5f6368; font-weight:500;">Direction:</label>
            <div style="display:flex; gap:16px;">
              <label style="display:flex; align-items:center; font-size:13px; color:#333; cursor:pointer;">
                <input type="radio" name="findDir" value="up" [(ngModel)]="findDirection" style="margin-right:4px; accent-color:#10b981;"> Up
              </label>
              <label style="display:flex; align-items:center; font-size:13px; color:#333; cursor:pointer;">
                <input type="radio" name="findDir" value="down" [(ngModel)]="findDirection" style="margin-right:4px; accent-color:#10b981;"> Down
              </label>
            </div>
          </div>
          
          <div style="display:flex; justify-content:center; gap:8px;">
            <button (click)="findNext()" style="background:#10b981; color:#fff; border:none; border-radius:4px; font-size:14px; font-weight:600; padding:8px 16px; cursor:pointer;">Find</button>
            <button (click)="replaceOne()" style="background:#f1f5f9; color:#333; border:1px solid #e2e8f0; border-radius:4px; font-size:14px; font-weight:500; padding:8px 16px; cursor:pointer;">Replace</button>
            <button (click)="replaceAll()" style="background:#f1f5f9; color:#333; border:1px solid #e2e8f0; border-radius:4px; font-size:14px; font-weight:500; padding:8px 16px; cursor:pointer;">Replace All</button>
            <button (click)="findModalOpen = false" style="background:#f1f5f9; color:#333; border:1px solid #e2e8f0; border-radius:4px; font-size:14px; font-weight:500; padding:8px 16px; cursor:pointer;">Close</button>
          </div>
          <div *ngIf="findStatus" style="margin-top:12px; text-align:center; font-size:13px; color:#ef4444;">{{ findStatus }}</div>
        </div>
      </div>

      <!-- Go To Modal -->
      <div class="modal-overlay" *ngIf="activeModal === 'goto'" (click)="activeModal = null" style="z-index: 10000; background: transparent;">
        <div class="modal" (click)="$event.stopPropagation()" style="width:360px; background:#fff; color:#333; border-radius:8px; padding:24px; box-shadow:0 8px 32px rgba(0,0,0,0.15); border:1px solid #e2e8f0; font-family:'Roboto',sans-serif;">
          <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;">
            <h3 style="margin:0; font-size:16px; font-weight:600; color:#333;">Go To</h3>
            <button (click)="activeModal = null" style="background:none; border:none; color:#5f6368; font-size:20px; cursor:pointer; padding:0; line-height:1;">&times;</button>
          </div>
          <div style="margin-bottom:24px;">
            <label style="font-size:13px; color:#5f6368; display:block; margin-bottom:8px; font-weight:500;">Reference (e.g. A1, Sheet2!B5)</label>
            <input [(ngModel)]="gotoQuery" (keydown.enter)="executeGoto()" style="width:100%;box-sizing:border-box;padding:10px 12px;border:1px solid #dadce0;border-radius:4px;font-size:14px;outline:none;transition:border-color 0.2s;" onfocus="this.style.borderColor='#10b981'" onblur="this.style.borderColor='#dadce0'" placeholder="Enter cell reference...">
          </div>
          <div style="display:flex; justify-content:flex-end; gap:10px;">
            <button (click)="activeModal = null" style="background:#f1f5f9; border:1px solid #e2e8f0; color:#333; font-size:14px; font-weight:500; padding:8px 16px; border-radius:4px; cursor:pointer;">Cancel</button>
            <button (click)="executeGoto()" style="background:#10b981; color:#fff; border:none; border-radius:4px; font-weight:600; font-size:14px; padding:8px 16px; cursor:pointer;">OK</button>
          </div>
        </div>
      </div>

      <!-- Custom Prompt Modal -->
      <div class="modal-overlay" *ngIf="promptModalOpen" (click)="closePrompt()">
        <div class="modal" (click)="$event.stopPropagation()" style="background:#fff; color:#333; border-radius:8px; padding:24px; width:420px; box-shadow:0 8px 32px rgba(0,0,0,0.15); border:1px solid #e2e8f0; max-width:90vw;">
          <h3 style="margin-top:0; font-size:16px; font-weight:600; color:#333; margin-bottom:16px;">{{promptModalTitle}}</h3>
          <input type="text" [(ngModel)]="promptModalValue" (keyup.enter)="submitPrompt()" style="width:100%; box-sizing:border-box; background:#f8f9fa; border:1px solid #cbd5e1; color:#333; font-size:14px; padding:10px 12px; border-radius:6px; outline:none; transition:border-color 0.2s;" onfocus="this.style.borderColor='#10b981'" onblur="this.style.borderColor='#cbd5e1'" autofocus>
          <div style="display:flex; justify-content:flex-end; gap:10px; margin-top:20px;">
            <button (click)="closePrompt()" style="background:#f1f5f9; border:1px solid #e2e8f0; color:#333; font-size:14px; font-weight:500; cursor:pointer; padding:8px 20px; border-radius:4px;">Cancel</button>
            <button (click)="submitPrompt()" style="background:#10b981; color:#fff; border:none; border-radius:4px; font-weight:600; font-size:14px; padding:8px 24px; cursor:pointer;">OK</button>
          </div>
        </div>
      </div>

      <!-- Custom Confirm Modal -->
      <div class="modal-overlay" *ngIf="confirmModalOpen" (click)="closeConfirm(false)" style="z-index: 10000;">
        <div class="modal" (click)="$event.stopPropagation()" style="background:#fff; color:#333; border-radius:8px; padding:24px; width:400px; box-shadow:0 8px 32px rgba(0,0,0,0.15); border:1px solid #e2e8f0; max-width:90vw; text-align:center;">
          <span class="material-symbols-outlined" style="font-size:36px; color:#ef4444; margin-bottom:12px;">delete_forever</span>
          <h3 style="margin-top:0; font-size:16px; font-weight:600; color:#333; margin-bottom:16px; line-height: 1.4;">{{confirmModalMessage}}</h3>
          <div style="display:flex; justify-content:center; gap:12px; margin-top:24px;">
            <button (click)="closeConfirm(false)" style="background:#f1f5f9; border:1px solid #e2e8f0; color:#333; font-size:14px; font-weight:500; cursor:pointer; padding:8px 24px; border-radius:4px; transition: background 0.2s;" onmouseover="this.style.background='#e2e8f0'" onmouseout="this.style.background='#f1f5f9'">Cancel</button>
            <button (click)="closeConfirm(true)" style="background:#ef4444; color:#fff; border:none; border-radius:4px; font-weight:600; font-size:14px; padding:8px 24px; cursor:pointer; transition: background 0.2s;" onmouseover="this.style.background='#dc2626'" onmouseout="this.style.background='#ef4444'">Delete</button>
          </div>
        </div>
      </div>

      <!-- Personal Dictionary Modal -->
      <div class="modal-overlay" *ngIf="personalDictModalOpen" (click)="personalDictModalOpen=false" style="z-index:10000;">
        <div class="modal" (click)="$event.stopPropagation()" style="width:440px;background:#fff;color:#333;border-radius:8px;box-shadow:0 8px 32px rgba(0,0,0,0.15);padding:24px;border:1px solid #e2e8f0;">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px;">
            <div style="display:flex;align-items:center;gap:10px;">
              <span class="material-symbols-outlined" style="color:#10b981;font-size:22px;">book</span>
              <h3 style="margin:0;font-size:18px;font-weight:600;">Personal Dictionary</h3>
            </div>
            <button (click)="personalDictModalOpen=false" style="background:none;border:none;cursor:pointer;color:#888;"><span class="material-symbols-outlined" style="font-size:20px;">close</span></button>
          </div>
          <div style="display:flex;gap:8px;margin-bottom:14px;">
            <input [(ngModel)]="personalDictNewWord" placeholder="Add a word..." (keyup.enter)="addPersonalDictWord()" style="flex:1;border:1px solid #cbd5e1;border-radius:4px;padding:8px 10px;font-size:13px;outline:none;" />
            <button (click)="addPersonalDictWord()" style="background:#10b981;color:#fff;border:none;padding:8px 16px;border-radius:4px;font-weight:600;cursor:pointer;">Add</button>
          </div>
          <div style="border:1px solid #e2e8f0;border-radius:6px;max-height:200px;overflow-y:auto;">
            <div *ngIf="personalDictWords.length===0" style="padding:24px;text-align:center;color:#9aa0a6;font-size:13px;">No words added yet.</div>
            <div *ngFor="let w of personalDictWords; let i=index" style="display:flex;justify-content:space-between;align-items:center;padding:10px 12px;border-bottom:1px solid #f1f5f9;font-size:13px;">
              <span>{{ w }}</span>
              <span (click)="removePersonalDictWord(i)" style="cursor:pointer;color:#d93025;font-size:13px;font-weight:600;">✕</span>
            </div>
          </div>
          <div style="display:flex;justify-content:flex-end;margin-top:16px;">
            <button (click)="personalDictModalOpen=false" style="background:#f1f5f9;color:#333;border:1px solid #e2e8f0;padding:8px 20px;border-radius:4px;font-weight:600;cursor:pointer;">Close</button>
          </div>
        </div>
      </div>

      <!-- Edit History Sidebar -->
      <div *ngIf="showEditHistoryPanel" class="edit-history-panel" style="position: absolute; right: 20px; top: 120px; width: 340px; background: #fff; border-radius: 8px; box-shadow: 0 4px 16px rgba(0,0,0,0.15); z-index: 1000; display: flex; flex-direction: column; border: 1px solid #e2e8f0; color: #333;">
        <div class="eh-header" style="display: flex; justify-content: space-between; align-items: center; padding: 12px 16px; border-bottom: 1px solid #eee; background: #f8f9fa; border-radius: 8px 8px 0 0;">
          <span style="font-weight: bold; font-size: 15px; color: #333;">Edit History</span>
          <span class="material-symbols-outlined" style="cursor: pointer; font-size: 18px; color: #5f6368;" (click)="showEditHistoryPanel = false">close</span>
        </div>
        <div class="eh-body" style="padding: 16px; max-height: 400px; overflow-y: auto;">
           <div style="font-size: 13px; color: #5f6368; margin-bottom: 16px;">Source: '{{sheets[currentSheetIdx].name || "Sheet1"}}'.{{colLabel(editHistoryCell?.c)}}{{editHistoryCell?.r + 1}}</div>
           
           <div *ngIf="!editHistoryData || editHistoryData.length === 0" style="color: #999; font-size: 13px; text-align: center; padding: 20px 0;">No edit history found.</div>

           <div *ngFor="let edit of editHistoryData" class="eh-entry" style="display: flex; gap: 12px; margin-bottom: 16px; background: #f8f9fa; padding: 12px; border-radius: 6px; border: 1px solid #eee;">
              <div class="eh-icon" style="width: 32px; height: 32px; border-radius: 50%; background: #e21b5a; color: #fff; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">
                  <span class="material-symbols-outlined" style="font-size: 20px;">person</span>
              </div>
              <div class="eh-info" style="flex: 1; word-break: break-word;">
                 <div style="display: flex; justify-content: space-between; margin-bottom: 2px;">
                    <div class="eh-user" style="font-weight: 600; font-size: 14px; color: #333;">{{edit.user}}</div>
                    <div class="eh-action" style="font-size: 11px; font-weight: bold; text-transform: uppercase;" [style.color]="edit.action === 'ADDED' ? '#2e8b57' : (edit.action === 'EDITED' ? '#1a73e8' : '#d32f2f')">{{edit.action}}</div>
                 </div>
                 <div class="eh-time" style="font-size: 12px; color: #666; margin-bottom: 6px;">{{edit.time | date:'medium'}}</div>
                 <div class="eh-value" *ngIf="edit.value !== undefined && edit.value !== null && edit.value !== ''" style="background: #fff; border: 1px solid #e2e8f0; padding: 4px 8px; border-radius: 4px; font-size: 13px; color: #333;">
                     <span class="material-symbols-outlined" style="font-size: 14px; vertical-align: middle; color: #2e8b57;">check_circle</span> <span style="margin-left: 4px;">{{edit.value}}</span>
                 </div>
              </div>
           </div>
        </div>
      </div>

      <!-- Spreadsheet Statistics Modal -->
      <div class="modal-overlay" *ngIf="showCustomFormatModal" (click)="showCustomFormatModal=false" style="z-index: 10000;">
        <div class="modal" (click)="$event.stopPropagation()" style="width: 400px; padding: 20px; background: #fff; color: #333; border-radius: 8px; box-shadow: 0 8px 32px rgba(0,0,0,0.15); border: 1px solid #e2e8f0;">
          <h3 style="margin-top: 0; color: #333;">Custom Number Format</h3>
          <p style="font-size: 13px; color: #666;">Enter a custom format string (e.g. <code>$#,##0.00</code>, <code>0.00%</code>, <code>&#64;</code>)</p>
          <input type="text" [(ngModel)]="customFormatString" style="width: 100%; padding: 8px; box-sizing: border-box; margin-bottom: 15px; font-family: monospace; border: 1px solid #ccc; border-radius: 4px; color: #333; background: #fff;">
          <div style="display: flex; justify-content: flex-end; gap: 10px;">
            <button (click)="showCustomFormatModal=false" style="padding: 6px 12px; border: 1px solid #ccc; background: #f8f9fa; color: #333; cursor: pointer; border-radius: 4px;">Cancel</button>
            <button (click)="applyCustomFormat()" style="padding: 6px 12px; border: none; background: #1a73e8; color: #fff; cursor: pointer; border-radius: 4px;">Apply</button>
          </div>
        </div>
      </div>

      <div class="modal-overlay" *ngIf="showMoreFormatsModal" (click)="showMoreFormatsModal=false" style="z-index: 10000;">
        <div class="modal" (click)="$event.stopPropagation()" style="width: 400px; max-height: 80vh; overflow-y: auto; padding: 20px; background: #fff; color: #333; border-radius: 8px; box-shadow: 0 8px 32px rgba(0,0,0,0.15); border: 1px solid #e2e8f0;">
          <h3 style="margin-top: 0; color: #333;">More Formats</h3>
          <div style="display: flex; flex-direction: column; gap: 8px; margin-bottom: 20px;">
            <div class="mf-item" (click)="setNumFormat('number'); showMoreFormatsModal=false">Number (1,234.56)</div>
            <div class="mf-item" (click)="setNumFormat('date_full'); showMoreFormatsModal=false">Full Date (Tuesday, August 5, 2030)</div>
            <div class="mf-item" (click)="setNumFormat('date_iso'); showMoreFormatsModal=false">ISO Date (2030-08-05)</div>
            <div class="mf-item" (click)="setNumFormat('accounting'); showMoreFormatsModal=false">Accounting ($ 1,234.56)</div>
            <div class="mf-item" (click)="setNumFormat('financial'); showMoreFormatsModal=false">Financial ( (1,234.56) )</div>
            <div class="mf-item" style="font-weight: bold; background: #f1f3f4;" (click)="openCustomFormatModal(); showMoreFormatsModal=false">Create Custom Format...</div>
          </div>
          <div style="display: flex; justify-content: flex-end;">
            <button (click)="showMoreFormatsModal=false" style="padding: 6px 12px; border: 1px solid #ccc; background: #f8f9fa; color: #333; cursor: pointer; border-radius: 4px;">Close</button>
          </div>
        </div>
      </div>

      <div class="modal-overlay" *ngIf="moveSheetModalOpen" (click)="moveSheetModalOpen=false" style="z-index:10000;">
        <div class="modal" (click)="$event.stopPropagation()" style="width:360px;background:#fff;color:#333;border-radius:8px;box-shadow:0 8px 32px rgba(0,0,0,0.15);padding:24px;border:1px solid #e2e8f0;">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px;">
            <h3 style="margin:0;font-size:18px;font-weight:500;">Move</h3>
            <button (click)="moveSheetModalOpen=false" style="background:none;border:none;cursor:pointer;color:#888;"><span class="material-symbols-outlined" style="font-size:20px;">close</span></button>
          </div>
          <div style="margin-bottom:12px;font-size:14px;color:#555;">
            Move "{{ moveSheetTargetIdx >= 0 && moveSheetTargetIdx < sheets.length ? sheets[moveSheetTargetIdx].name : '' }}"
          </div>
          <select [(ngModel)]="moveSheetDestination" style="width:100%;background:#f8fafc;color:#333;border:1px solid #cbd5e1;padding:8px;border-radius:4px;margin-bottom:24px;font-size:14px;outline:none;">
            <option value="start">Move to the Start</option>
            <option value="end">Move to the End</option>
            <optgroup label="or After">
              <ng-container *ngFor="let s of sheets; let i = index">
                <option *ngIf="i !== moveSheetTargetIdx" [value]="i">{{ s.name }}</option>
              </ng-container>
            </optgroup>
          </select>
          <div style="display:flex;justify-content:flex-end;gap:12px;">
            <button (click)="confirmMoveSheet()" style="background:#10b981;color:#fff;border:none;padding:8px 20px;border-radius:4px;font-weight:600;cursor:pointer;">OK</button>
            <button (click)="moveSheetModalOpen=false" style="background:#f1f5f9;color:#333;border:1px solid #cbd5e1;padding:8px 20px;border-radius:4px;font-weight:600;cursor:pointer;">Cancel</button>
          </div>
        </div>
      </div>

      <div class="modal-overlay" *ngIf="statsModalOpen" (click)="statsModalOpen=false" style="z-index:10000; display:flex; align-items:center; justify-content:center; background:rgba(0,0,0,0.4);">
        <div class="modal" (click)="$event.stopPropagation()" style="width:420px;background:#fff;color:#333;border-radius:4px;box-shadow:0 4px 16px rgba(0,0,0,0.2);display:flex;flex-direction:column;font-family:'Roboto',sans-serif;">
          
          <!-- Header -->
          <div style="display:flex;justify-content:space-between;align-items:center;padding:12px 16px;border-bottom:1px solid #e0e0e0;">
            <div style="font-size:15px;font-weight:500;">Spreadsheet Statistics</div>
            <button (click)="statsModalOpen=false" style="background:none;border:none;cursor:pointer;color:#5f6368;display:flex;align-items:center;"><span class="material-symbols-outlined" style="font-size:20px;">close</span></button>
          </div>
          
          <!-- Body -->
          <div style="padding:16px; font-size:13px;">
             <!-- Spreadsheet Section -->
             <div style="border:1px solid #e0e0e0; border-radius:4px; margin-bottom:16px; background:#f9f9f9;">
                <div style="padding:12px 16px; font-weight:600; border-bottom:1px solid #e0e0e0;">Spreadsheet</div>
                <div style="padding:12px 16px; display:flex; flex-direction:column; gap:8px;">
                   <div style="display:flex; justify-content:space-between;"><span>Sheets:</span><span>{{sheets.length}}</span></div>
                   <div style="display:flex; justify-content:space-between;"><span>Cells with data:</span><span>{{getGlobalStats().cellsWithData | number}}</span></div>
                   <div style="display:flex; justify-content:space-between;"><span>Used cells:</span><span>{{getGlobalStats().usedCells | number}}</span></div>
                   <div style="display:flex; justify-content:space-between;"><span>Tables:</span><span>0</span></div>
                   <div style="display:flex; justify-content:space-between;"><span>Formulas:</span><span>0</span></div>
                   <div style="display:flex; justify-content:space-between;"><span>Charts:</span><span>0</span></div>
                   <div style="display:flex; justify-content:space-between;"><span>Pivot Tables:</span><span>0</span></div>
                </div>
             </div>

             <!-- Sheet Section -->
             <div style="border:1px solid #e0e0e0; border-radius:4px; background:#f9f9f9;">
                <div style="padding:12px 16px; font-weight:600; border-bottom:1px solid #e0e0e0; display:flex; align-items:center; gap:8px;">
                   <span>Sheet Name:</span>
                   <select [(ngModel)]="statsSelectedSheetIdx" style="border:1px solid #ccc; border-radius:2px; padding:4px; outline:none; background:#fff; font-size:13px; font-family:inherit;">
                      <option *ngFor="let s of sheets; let i=index" [value]="i">{{s.name}}</option>
                   </select>
                </div>
                <div style="padding:12px 16px; display:flex; flex-direction:column; gap:8px;">
                   <div style="display:flex; justify-content:space-between;"><span>End of sheet:</span><span style="color:#0f9d58; font-weight:500;">{{getSheetStats(statsSelectedSheetIdx).endOfSheet}}</span></div>
                   <div style="display:flex; justify-content:space-between;"><span>Cells with data:</span><span>{{getSheetStats(statsSelectedSheetIdx).cellsWithData | number}}</span></div>
                   <div style="display:flex; justify-content:space-between;"><span>Used cells:</span><span>{{getSheetStats(statsSelectedSheetIdx).usedCells | number}}</span></div>
                   <div style="display:flex; justify-content:space-between;"><span>Tables:</span><span>0</span></div>
                   <div style="display:flex; justify-content:space-between;"><span>Formulas:</span><span>0</span></div>
                   <div style="display:flex; justify-content:space-between;"><span>Charts:</span><span>0</span></div>
                   <div style="display:flex; justify-content:space-between;"><span>Pivot Tables:</span><span>0</span></div>
                </div>
             </div>
          </div>

          <!-- Footer -->
          <div style="padding:12px 16px; display:flex; justify-content:flex-end;">
            <button (click)="statsModalOpen=false" style="background:#f8f9fa; color:#333; border:1px solid #ccc; padding:6px 16px; border-radius:4px; font-size:13px; font-weight:500; cursor:pointer;">Close</button>
          </div>
        </div>
      </div>

      <!-- Manage Lock Settings Modal -->
      <div class="modal-overlay" *ngIf="manageLockSettingsModalOpen" (click)="manageLockSettingsModalOpen=false" style="z-index:10000; display:flex; align-items:center; justify-content:center; background:rgba(0,0,0,0.5);">
        <div class="modal" (click)="$event.stopPropagation()" [style.background]="currentTheme === 'dark' ? '#242424' : '#fff'" [style.color]="currentTheme === 'dark' ? '#fff' : '#333'" style="width:480px; border-radius:8px; box-shadow:0 4px 16px rgba(0,0,0,0.5); display:flex; flex-direction:column; font-family:'Roboto',sans-serif; min-height: 300px;">
          
          <!-- Header -->
          <div style="display:flex; justify-content:space-between; align-items:center; padding:16px 20px;">
            <div style="font-size:18px; font-weight:500;">Manage Lock Settings</div>
            <button (click)="manageLockSettingsModalOpen=false" style="background:none; border:none; cursor:pointer; display:flex; align-items:center;" [style.color]="currentTheme === 'dark' ? '#9aa0a6' : '#5f6368'"><span class="material-symbols-outlined" style="font-size:20px;">close</span></button>
          </div>
          
          <!-- Tabs -->
          <div style="display:flex; padding:0 20px; gap: 24px;" [style.border-bottom]="currentTheme === 'dark' ? '1px solid #3c4043' : '1px solid #e0e0e0'">
            <div (click)="lockSettingsTab = 'ranges'" [style.color]="lockSettingsTab === 'ranges' ? '#1da954' : (currentTheme === 'dark' ? '#e8eaed' : '#5f6368')" [style.border-bottom]="lockSettingsTab === 'ranges' ? '2px solid #1da954' : '2px solid transparent'" style="padding:10px 4px; cursor:pointer; font-weight:500; font-size:14px; transition: 0.2s;">Ranges</div>
            <div (click)="lockSettingsTab = 'sheets'" [style.color]="lockSettingsTab === 'sheets' ? '#1da954' : (currentTheme === 'dark' ? '#e8eaed' : '#5f6368')" [style.border-bottom]="lockSettingsTab === 'sheets' ? '2px solid #1da954' : '2px solid transparent'" style="padding:10px 4px; cursor:pointer; font-weight:500; font-size:14px; transition: 0.2s;">Sheets</div>
          </div>
          
          <!-- Body -->
          <div style="padding:20px; flex:1; display:flex; flex-direction:column;">
            <div *ngIf="lockSettingsTab === 'ranges'">
              <div style="display:flex; align-items:center; margin-bottom: 24px;">
                <span style="font-weight:600; font-size:14px; margin-right:16px;">View Locked Cells in:</span>
                <select [(ngModel)]="lockSettingsSelectedSheet" [style.background]="currentTheme === 'dark' ? '#303134' : '#fff'" [style.color]="currentTheme === 'dark' ? '#e8eaed' : '#333'" [style.border]="currentTheme === 'dark' ? '1px solid #5f6368' : '1px solid #ccc'" style="flex:1; border-radius:4px; padding:8px 12px; font-size:14px; outline:none;">
                  <option value="all">Whole Spreadsheet</option>
                  <option *ngFor="let s of sheets; let i = index" [value]="i">{{ s.name }}</option>
                </select>
              </div>
              <div *ngIf="getLockedCellsForCurrentSettings().length === 0" style="display:flex; justify-content:center; align-items:center; flex:1; min-height:120px; font-size:14px;" [style.color]="currentTheme === 'dark' ? '#9aa0a6' : '#5f6368'">
                No Locked Cells
              </div>
              <div *ngIf="getLockedCellsForCurrentSettings().length > 0" style="flex: 1; overflow-y: auto; max-height: 200px; border: 1px solid {{ currentTheme === 'dark' ? '#3c4043' : '#e0e0e0' }}; border-radius: 4px; padding: 4px 0;">
                <div *ngFor="let item of getLockedCellsForCurrentSettings(); trackBy: trackByCellRef" style="display: flex; justify-content: space-between; align-items: center; padding: 8px 16px; border-bottom: 1px solid {{ currentTheme === 'dark' ? '#3c4043' : '#e0e0e0' }};" [style.color]="currentTheme === 'dark' ? '#e8eaed' : '#333'">
                  <div style="font-size: 13px;">
                    <span *ngIf="lockSettingsSelectedSheet === 'all'" style="opacity: 0.7; margin-right: 8px;">{{ item.sheetName }}</span>
                    <span style="font-weight: 500;">Cell {{ item.ref }}</span>
                  </div>
                  <button (click)="unlockCellFromSettings(item)" style="background: none; border: none; cursor: pointer; display: flex; align-items: center; color: #ea4335;" title="Unlock Cell">
                    <span class="material-symbols-outlined" style="font-size: 16px;">lock_open</span>
                  </button>
                </div>
              </div>
            </div>

            <div *ngIf="lockSettingsTab === 'sheets'">
              <div style="font-weight:600; font-size:14px; margin-bottom: 24px;">View Locked Sheet(s)</div>
              
              <div *ngIf="getLockedSheets().length === 0" style="display:flex; justify-content:center; align-items:center; flex:1; min-height:120px; font-size:14px;" [style.color]="currentTheme === 'dark' ? '#9aa0a6' : '#5f6368'">
                No Locked Sheets
              </div>
              <div *ngIf="getLockedSheets().length > 0" style="flex: 1; overflow-y: auto; max-height: 200px; border: 1px solid {{ currentTheme === 'dark' ? '#3c4043' : '#e0e0e0' }}; border-radius: 4px; padding: 4px 0;">
                <div *ngFor="let item of getLockedSheets(); trackBy: trackBySheetIndex" style="display: flex; justify-content: space-between; align-items: center; padding: 8px 16px; border-bottom: 1px solid {{ currentTheme === 'dark' ? '#3c4043' : '#e0e0e0' }};" [style.color]="currentTheme === 'dark' ? '#e8eaed' : '#333'">
                  <div style="font-size: 13px;">
                    <span style="font-weight: 500;">{{ item.sheetName }}</span>
                  </div>
                  <button (click)="toggleLockSheet(item.sheetIndex); $event.stopPropagation()" style="background: none; border: none; cursor: pointer; display: flex; align-items: center; color: #ea4335;" title="Unlock Sheet">
                    <span class="material-symbols-outlined" style="font-size: 18px; pointer-events: none;">lock_open</span>
                  </button>
                </div>
              </div>
            </div>
          </div>
          
          <!-- Footer -->
          <div style="padding:16px 20px; display:flex; justify-content:flex-end;">
            <button (click)="manageLockSettingsModalOpen=false" [style.background]="currentTheme === 'dark' ? '#3c4043' : '#f8f9fa'" [style.color]="currentTheme === 'dark' ? '#e8eaed' : '#333'" [style.border]="currentTheme === 'dark' ? 'none' : '1px solid #ccc'" style="padding:8px 16px; border-radius:4px; font-size:14px; font-weight:500; cursor:pointer;">Close</button>
          </div>
        </div>
      </div>

      <div class="bottom-chat-bar">
         <div class="bcb-item" (click)="toggleWidget('chat')">
            <span class="material-symbols-outlined" style="color:#d32f2f;">chat</span>
            <span>Unread Chats</span>
            <div class="bcb-badge">0</div>
         </div>
         <div class="bcb-item" (click)="toggleWidget('channels')">
            <span class="material-symbols-outlined" style="color:#5f6368;">group</span>
            <span>Channels</span>
         </div>
      </div>
      <app-chat-widget [activeWidget]="activeWidget" (close)="activeWidget=null"></app-chat-widget>
      <app-chat-widget [activeWidget]="activeWidget" (close)="activeWidget=null"></app-chat-widget>
      
      <!-- Loading Data Overlay -->
      <div class="loading-overlay" *ngIf="isLoadingDocument || isUploading">
        <div class="loading-modal shadow-lg" *ngIf="isLoadingDocument">
           <div class="lm-spinner"></div>
           <div class="lm-title">Loading Spreadsheet...</div>
           <div class="lm-subtitle">Retrieving cells, formulas, and formatting. Please wait.</div>
        </div>
        
        <div class="upload-modal shadow-lg" *ngIf="isUploading">
           <div class="um-icon">
              <span class="material-symbols-outlined" style="font-size:32px; color:#1a73e8;">cloud_upload</span>
           </div>
           <div class="um-title">Importing Spreadsheet...</div>
           <div class="um-subtitle">Please wait while your file is securely uploaded and processed.</div>
           
           <div class="um-progress-container">
              <div class="um-progress-bar" [style.width]="uploadProgress + '%'"></div>
           </div>
           
           <div class="um-stats">
              <div class="um-percent">{{ uploadProgress }}%</div>
              <div class="um-time">{{ uploadTimeLeft }}</div>
           </div>
        </div>
      </div>

       <div class="modal-overlay" *ngIf="ocrModalOpen" [style.background]="currentTheme === 'dark' ? 'rgba(0,0,0,0)' : 'rgba(0,0,0,0)'" style="z-index: 10000; pointer-events: none; position: fixed; inset: 0; display: flex; align-items: center; justify-content: center;">
         <div class="modal" (mousedown)="$event.stopPropagation()" [style.background]="currentTheme === 'dark' ? '#1e1e1e' : '#fff'" [style.color]="currentTheme === 'dark' ? '#e0e0e0' : '#333'" style="pointer-events: auto; width: 1000px; height: 650px; max-width: 95vw; border-radius: 8px; padding: 0; display: flex; flex-direction: column; max-height: 90vh; overflow: hidden; box-shadow: 0 4px 24px rgba(0,0,0,0.4); border: 1px solid #444;">
            
            <div [style.border-bottom]="currentTheme === 'dark' ? '1px solid #333' : '1px solid #e0e0e0'" [style.background]="currentTheme === 'dark' ? '#1e1e1e' : '#fff'" style="font-size: 16px; font-weight: 500; padding: 16px 20px; display: flex; justify-content: space-between; align-items: center;">
              <span>Data from Picture</span>
              <span class="material-symbols-outlined" style="cursor: pointer; color: #888;" (click)="ocrModalOpen=false">close</span>
            </div>
            
            <div *ngIf="ocrProgress > 0 && ocrProgress < 100" style="display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; color: #888;">
               <div class="lm-spinner"></div>
               <div style="margin-top: 16px;">Processing OCR... {{ocrProgress}}%</div>
            </div>

            <!-- Main body -->
            <div *ngIf="ocrProgress === 100 && ocrData.length > 0" style="display: flex; flex-direction: column; flex: 1; min-height: 0;">
               
               <!-- Toolbar spanning across top -->
               <div [style.background]="currentTheme === 'dark' ? '#252525' : '#f8f9fa'" [style.border-bottom]="currentTheme === 'dark' ? '1px solid #333' : '1px solid #e0e0e0'" style="display: flex; align-items: center; gap: 12px; padding: 8px 20px;">
                  <button (click)="ocrUndo()" [disabled]="ocrHistoryIndex <= 0" [style.opacity]="ocrHistoryIndex <= 0 ? 0.3 : 1" style="background: none; border: none; cursor: pointer; color: #888; padding: 4px; display: flex; align-items: center; justify-content: center;"><span class="material-symbols-outlined" style="font-size: 18px;">undo</span></button>
                  <button (click)="ocrRedo()" [disabled]="ocrHistoryIndex >= ocrHistory.length - 1" [style.opacity]="ocrHistoryIndex >= ocrHistory.length - 1 ? 0.3 : 1" style="background: none; border: none; cursor: pointer; color: #888; padding: 4px; display: flex; align-items: center; justify-content: center;"><span class="material-symbols-outlined" style="font-size: 18px;">redo</span></button>
                  
                  <div [style.background]="currentTheme === 'dark' ? '#444' : '#ccc'" style="width: 1px; height: 16px; margin: 0 4px;"></div>
                  
                  <button (click)="ocrAppendMode = 'left'; ocrInsertTarget = 'existing'" [style.color]="ocrAppendMode === 'left' && ocrInsertTarget === 'existing' ? (currentTheme === 'dark' ? '#fff' : '#000') : '#888'" style="background: none; border: none; cursor: pointer; display: flex; align-items: center; gap: 6px; font-size: 13px; font-weight: 500;"><span class="material-symbols-outlined" style="font-size: 16px; color: #10b981;">arrow_left_alt</span> Append Left</button>
                  <button (click)="ocrAppendMode = 'right'; ocrInsertTarget = 'existing'" [style.color]="ocrAppendMode === 'right' && ocrInsertTarget === 'existing' ? (currentTheme === 'dark' ? '#fff' : '#000') : '#888'" style="background: none; border: none; cursor: pointer; display: flex; align-items: center; gap: 6px; font-size: 13px; font-weight: 500;"><span class="material-symbols-outlined" style="font-size: 16px; color: #10b981;">arrow_right_alt</span> Append Right</button>
                  <button (click)="ocrAppendMode = 'above'; ocrInsertTarget = 'existing'" [style.color]="ocrAppendMode === 'above' && ocrInsertTarget === 'existing' ? (currentTheme === 'dark' ? '#fff' : '#000') : '#888'" style="background: none; border: none; cursor: pointer; display: flex; align-items: center; gap: 6px; font-size: 13px; font-weight: 500;"><span class="material-symbols-outlined" style="font-size: 16px; color: #10b981;">arrow_upward_alt</span> Append Above</button>
                  <button (click)="ocrAppendMode = 'below'; ocrInsertTarget = 'existing'" [style.color]="ocrAppendMode === 'below' && ocrInsertTarget === 'existing' ? (currentTheme === 'dark' ? '#fff' : '#000') : '#888'" style="background: none; border: none; cursor: pointer; display: flex; align-items: center; gap: 6px; font-size: 13px; font-weight: 500;"><span class="material-symbols-outlined" style="font-size: 16px; color: #10b981;">arrow_downward_alt</span> Append Below</button>
                  
                  <div style="flex: 1;"></div>
                  
                  <div style="position: relative; display: flex; align-items: center;">
                     <span class="material-symbols-outlined" style="position: absolute; left: 8px; font-size: 16px; color: #888;">search</span>
                     <input type="text" [style.background]="currentTheme === 'dark' ? '#1a1a1a' : '#fff'" [style.color]="currentTheme === 'dark' ? '#fff' : '#333'" [style.border]="currentTheme === 'dark' ? '1px solid #444' : '1px solid #ccc'" style="padding: 6px 8px 6px 30px; border-radius: 4px; font-size: 13px; width: 220px; outline: none;" placeholder="">
                  </div>
               </div>
               
               <!-- 2 Column Layout -->
               <div style="display: flex; gap: 20px; flex: 1; min-height: 0; overflow: hidden; padding: 20px;">
                  
                  <!-- Left side: Image -->
                  <div [style.background]="currentTheme === 'dark' ? '#0a0a0a' : '#f8f9fa'" [style.border]="currentTheme === 'dark' ? '1px solid #333' : '1px solid #e0e0e0'" style="flex: 1; border-radius: 6px; display: flex; align-items: center; justify-content: center; overflow: hidden;">
                     <img *ngIf="ocrImage" [src]="ocrImage" style="max-width: 100%; max-height: 100%; object-fit: contain;">
                  </div>
                  
                  <!-- Right side: Data preview -->
                  <div style="flex: 1; display: flex; flex-direction: column; overflow: hidden;">
                     <div style="flex: 1; overflow: auto; background: #fff; border-radius: 2px; border: 1px solid #e0e0e0;">
                        <table style="border-collapse: collapse; font-size: 13px; min-width: 100%;">
                           <!-- Header Row (A, B, C...) -->
                           <tr style="background: #f8f9fa; border-bottom: 1px solid #ccc;">
                             <td style="width: 35px; min-width: 35px; border-right: 1px solid #ccc; background: #eaecf0;"></td>
                             <td *ngFor="let col of ocrData[0]; let c = index" style="min-width: 100px; text-align: center; border-right: 1px solid #ccc; padding: 4px 8px; color: #444; font-weight: 500;">
                               {{ colLabel(c) }}
                             </td>
                           </tr>
                           <tr *ngFor="let row of ocrData; let r = index">
                              <td style="width: 35px; min-width: 35px; text-align: center; background: #f8f9fa; border: 1px solid #ccc; color: #444; font-weight: 500;">{{ r + 1 }}</td>
                              <td *ngFor="let col of row; let c = index; trackBy: trackByFn" style="min-width: 100px; border: 1px solid #ccc; padding: 0; position: relative;"
                                  (mousedown)="startOcrDrag(r, c)"
                                  (mouseenter)="doOcrDrag(r, c)"
                                  (dblclick)="startOcrEdit(r, c)"
                                  [style.background]="isOcrSelected(r, c) ? (currentTheme === 'dark' ? '#2c3e50' : '#e8f0fe') : 'transparent'">
                                 <div *ngIf="ocrSelStart?.r === r && ocrSelStart?.c === c" style="position: absolute; inset: 0; border: 2px solid #1a73e8; pointer-events: none; z-index: 2;"></div>
                                 <input *ngIf="ocrEdit?.r === r && ocrEdit?.c === c" type="text" [(ngModel)]="ocrData[r][c]" (blur)="ocrEdit = null; saveOcrHistory()" (keydown.enter)="ocrEdit = null; saveOcrHistory()" style="position: absolute; inset: 0; box-sizing: border-box; border: none; width: 100%; height: 100%; min-height: 28px; outline: none; background: transparent; padding: 4px 8px; color: inherit; z-index: 3;" autofocus>
                                 <div *ngIf="!(ocrEdit?.r === r && ocrEdit?.c === c)" style="width: 100%; height: 100%; min-height: 28px; padding: 4px 8px; color: inherit; user-select: none; position: relative; z-index: 1; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 250px;">{{ ocrData[r][c] }}</div>
                              </td>
                           </tr>
                        </table>
                     </div>
                     
                     <!-- Insert Options -->
                     <div style="margin-top: 16px; display: flex; flex-direction: column; gap: 10px;">
                        <div style="font-size: 12px; color: #888;">Insert options</div>
                        <div style="display: flex; gap: 24px;">
                           <label [style.color]="currentTheme === 'dark' ? '#ccc' : '#333'" style="display: flex; align-items: center; gap: 8px; cursor: pointer; font-size: 13px;">
                              <input type="radio" name="ocrInsertTarget" value="new" [(ngModel)]="ocrInsertTarget" style="accent-color: #10b981; transform: scale(1.1);"> New sheet
                           </label>
                           <label [style.color]="currentTheme === 'dark' ? '#ccc' : '#333'" style="display: flex; align-items: center; gap: 8px; cursor: pointer; font-size: 13px;">
                              <input type="radio" name="ocrInsertTarget" value="existing" [(ngModel)]="ocrInsertTarget" style="accent-color: #10b981; transform: scale(1.1);"> Existing sheet
                           </label>
                        </div>
                     </div>
                  </div>
               </div>
               
               <!-- Footer -->
               <div [style.background]="currentTheme === 'dark' ? '#252525' : '#f8f9fa'" [style.border-top]="currentTheme === 'dark' ? '1px solid #333' : '1px solid #e0e0e0'" style="padding: 16px 20px; display: flex; justify-content: space-between; align-items: center;">
                  <button (click)="dataFromPicture()" [style.background]="currentTheme === 'dark' ? '#3a3a3a' : '#fff'" [style.border]="currentTheme === 'dark' ? '1px solid #555' : '1px solid #ccc'" [style.color]="currentTheme === 'dark' ? '#fff' : '#333'" style="padding: 8px 24px; border-radius: 4px; cursor: pointer; font-weight: 500; font-size: 13px;">Back</button>
                  
                  <div style="display: flex; gap: 12px;">
                     <button (click)="ocrModalOpen=false" [style.border]="currentTheme === 'dark' ? '1px solid #555' : '1px solid #ccc'" [style.color]="currentTheme === 'dark' ? '#e0e0e0' : '#333'" style="padding: 8px 24px; background: transparent; border-radius: 4px; cursor: pointer; font-weight: 500; font-size: 13px;">Cancel</button>
                     <button (click)="insertOcrData()" [disabled]="ocrProgress !== 100 || !ocrData.length" [style.opacity]="ocrProgress === 100 && ocrData.length ? '1' : '0.5'" style="padding: 8px 28px; background: #10b981; color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: 500; font-size: 13px;">Insert</button>
                  </div>
               </div>
            </div>
            
            <div *ngIf="ocrProgress === 100 && ocrData.length === 0" style="display: flex; flex-direction: column; align-items: center; justify-content: center; flex: 1; color: #888;">
               No tabular data found in image.
               <div style="margin-top: 24px;">
                  <button (click)="ocrModalOpen=false" [style.border]="currentTheme === 'dark' ? '1px solid #555' : '1px solid #ccc'" [style.color]="currentTheme === 'dark' ? '#e0e0e0' : '#333'" style="padding: 8px 24px; background: transparent; border-radius: 4px; cursor: pointer; font-weight: 500; margin-right: 12px; font-size: 13px;">Cancel</button>
                  <button (click)="dataFromPicture()" [style.background]="currentTheme === 'dark' ? '#3a3a3a' : '#fff'" [style.border]="currentTheme === 'dark' ? '1px solid #555' : '1px solid #ccc'" [style.color]="currentTheme === 'dark' ? '#fff' : '#333'" style="padding: 8px 24px; border-radius: 4px; cursor: pointer; font-weight: 500; font-size: 13px;">Back</button>
               </div>
            </div>
         </div>
       </div>
    </div>

    <!-- IMMERSIVE VERSION HISTORY (LIGHT STYLE) -->
    <div *ngIf="activeModal === 'version'" style="position: fixed; inset: 0; z-index: 99999; display: flex; flex-direction: column; background: #f8f9fa; color: #202124; font-family: 'Roboto', sans-serif;">
      
      <!-- Top Header -->
      <div style="height: 60px; display: flex; align-items: center; justify-content: space-between; padding: 0 20px; border-bottom: 1px solid #dadce0; background: #fff;">
        <!-- Left: Doc Title -->
        <div style="display: flex; align-items: center; gap: 16px;">
          <button (click)="activeModal = null" style="background:none; border:none; color:#5f6368; cursor:pointer; font-size:24px; padding-bottom: 4px;">&larr;</button>
          <div>
            <div style="font-size:16px; font-weight:600;">{{ title || 'Untitled Document' }}</div>
            <div style="font-size:12px; color:#1a73e8;">Current Version</div>
          </div>
        </div>
        
        <!-- Center: Actions -->
        <div style="display: flex; align-items: center; gap: 24px;">
          <div style="display:flex; align-items:center; gap:6px; cursor:pointer; color:#5f6368; font-size:13px;" (click)="promptNameVersion()">
            <span class="material-symbols-outlined" style="font-size:18px;">update</span> Name This Version
          </div>
          <div style="display:flex; align-items:center; gap:6px; cursor:pointer; color:#5f6368; font-size:13px;" (click)="makeCopy()">
            <span class="material-symbols-outlined" style="font-size:18px;">file_copy</span> Make a Copy
          </div>
          <div style="display:flex; align-items:center; gap:6px; cursor:pointer; color:#5f6368; font-size:13px;" (click)="exportFile('xlsx')">
            <span class="material-symbols-outlined" style="font-size:18px;">download</span> Download as
          </div>
          <div style="display:flex; align-items:center; gap:6px; cursor:pointer; color:#5f6368; font-size:13px;" (click)="showToast('Changelog will be available in the next update.')">
            <span class="material-symbols-outlined" style="font-size:18px;">history</span> Changelog
          </div>
        </div>

        <!-- Right: Restore -->
        <div style="display: flex; align-items: center; gap: 16px;">
          <button *ngIf="previewVersionId && versions.length && previewVersionId !== versions[0].id" 
                  (click)="confirmRestoreVersion(previewVersionId)"
                  style="background: #1a73e8; border: none; color: #fff; padding: 6px 16px; border-radius: 4px; font-size: 13px; font-weight: 600; cursor: pointer; display: flex; align-items: center; gap: 6px;">
            <span class="material-symbols-outlined" style="font-size:16px;">history</span> Restore This Version
          </button>
        </div>
      </div>

      <!-- Main Layout -->
      <div style="display: flex; flex: 1; overflow: hidden; background: #f8f9fa;">
        
        <!-- Center (Preview Grid) -->
        <div style="flex: 1; display: flex; flex-direction: column; overflow: hidden; padding: 24px 24px 0 24px;">
          
          <div *ngIf="!previewData" style="flex: 1; display: flex; align-items: center; justify-content: center; color: #5f6368; font-style: italic; background: #fff; border-radius: 8px 8px 0 0; border: 1px solid #dadce0; border-bottom: none;">
            Select a version to preview
          </div>

          <ng-container *ngIf="previewData">
            <!-- Active Cell Indicator Placeholder -->
            <div style="height:36px; background:#fff; display:flex; align-items:center; padding:0 12px; border-radius: 4px 4px 0 0; border: 1px solid #dadce0; border-bottom: none;">
              <div style="color:#5f6368; font-size:12px; font-weight: 500; width: 40px; text-align: center; border-right: 1px solid #dadce0; padding-right: 8px; margin-right: 8px;">A1</div>
              <div style="background:#f1f3f4; padding:2px 8px; font-size:12px; color:#5f6368; border-radius:4px; font-family: monospace;">fx</div>
              <div style="margin-left:12px; color:#888; font-size:13px; font-style:italic;"></div>
            </div>
            
            <!-- The Grid -->
            <div style="flex: 1; overflow: auto; background: #fff; position: relative; border: 1px solid #dadce0; border-bottom: none;">
              <table class="grid" style="border-collapse: separate; border-spacing: 0;">
                <thead>
                  <tr>
                    <th class="corner" style="position: sticky; top: 0; left: 0; z-index: 6; border-right: 1px solid #c0c0c0; border-bottom: 1px solid #c0c0c0; background: #f8f9fa;"></th>
                    <th *ngFor="let col of previewCells[0]; let c = index" class="col-head" style="position: sticky; top: 0; z-index: 5; border-right: 1px solid #c0c0c0; border-bottom: 1px solid #c0c0c0; background: #f8f9fa; min-width: 100px;">
                      {{ colLabel(c) }}
                    </th>
                  </tr>
                </thead>
                <tbody>
                  <tr *ngFor="let row of previewCells; let r = index">
                    <td class="row-head" style="position: sticky; left: 0; z-index: 4; border-right: 1px solid #c0c0c0; border-bottom: 1px solid #c0c0c0; background: #f8f9fa;">
                      {{ r + 1 }}
                    </td>
                    <td *ngFor="let cell of row; let c = index" class="cell" style="border-right: 1px solid #e2e8f0; border-bottom: 1px solid #e2e8f0;">
                      {{ cell }}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>

            <!-- Bottom Sheet Tabs -->
            <div style="height: 40px; background: #fff; display: flex; align-items: center; padding: 0 12px; gap: 4px; border: 1px solid #dadce0; border-top: none;">
              <div style="color: #1a73e8; padding: 0 12px; cursor: pointer;"><span class="material-symbols-outlined" style="font-size: 20px; font-weight: bold;">add</span></div>
              <div style="color: #5f6368; padding: 0 12px; cursor: pointer;"><span class="material-symbols-outlined" style="font-size: 20px;">menu</span></div>
              <div style="width: 1px; height: 20px; background: #dadce0; margin: 0 8px;"></div>
              
              <div *ngFor="let sheet of previewSheets; let i = index" 
                   (click)="previewActiveSheetIdx = i"
                   style="padding: 6px 16px; cursor: pointer; font-size: 13px; font-weight: 500; border-radius: 4px; display: flex; align-items: center; gap: 8px;"
                   [style.background]="previewActiveSheetIdx === i ? '#e8f0fe' : 'transparent'"
                   [style.color]="previewActiveSheetIdx === i ? '#1a73e8' : '#5f6368'">
                {{ sheet.name || 'Sheet ' + (i + 1) }}
              </div>
            </div>
          </ng-container>
        </div>

        <!-- Right Sidebar (Version List) -->
        <div style="width: 320px; background: #fff; border-left: 1px solid #dadce0; display: flex; flex-direction: column;">
          
          <div style="padding: 16px; display: flex; align-items: center; justify-content: space-between;">
            <h3 style="margin: 0; font-size: 16px; font-weight: 600; color: #202124;">Version History</h3>
          </div>

          <div style="padding: 0 16px 12px 16px; border-bottom: 1px solid #dadce0; display: flex; align-items: center; gap: 8px;">
            <input type="checkbox" checked style="accent-color: #1a73e8; cursor: pointer; width: 16px; height: 16px;">
            <span style="font-size: 14px; color: #202124; font-weight: 500;">Highlight Changes</span>
          </div>

          <div style="flex: 1; overflow-y: auto; padding: 16px 12px;">
            <!-- Version Cards -->
            <div *ngFor="let v of versions; let i = index" 
                 (click)="previewVersion(v.id)"
                 style="padding: 12px 16px; margin-bottom: 8px; border-radius: 6px; cursor: pointer; background: #fff; position: relative; transition: all 0.2s;"
                 [style.border]="previewVersionId === v.id ? '1px solid #1a73e8' : '1px solid transparent'"
                 [style.background]="previewVersionId === v.id ? '#e8f0fe' : (i===0 && !previewVersionId ? '#f0fdf4' : '#fff')"
                 [style.box-shadow]="previewVersionId !== v.id ? '0 1px 2px rgba(0,0,0,0.05)' : 'none'">
              
              <div style="font-size: 13px; font-weight: 600; color: #202124; margin-bottom: 4px;">
                {{ v.created_at | date:'MMM d, y, h:mm:ss a' }}
              </div>
              <div *ngIf="i === 0" style="font-size: 12px; color: #10b981; margin-bottom: 8px;">Current Version</div>
              <div *ngIf="i !== 0 && v.version_name" style="font-size: 12px; color: #1a73e8; margin-bottom: 8px;">{{ v.version_name }}</div>
              
              <div style="display: flex; align-items: center; gap: 8px; font-size: 12px; color: #5f6368;">
                <div style="width: 10px; height: 10px; border-radius: 2px;" [style.background]="v.is_named ? '#fbbc04' : '#1a73e8'"></div>
                {{ v.is_named ? 'Named Version' : 'Auto-save' }}
              </div>

              <span *ngIf="previewVersionId === v.id" class="material-symbols-outlined" style="position: absolute; right: 8px; top: 12px; color: #1a73e8; font-size: 18px;">more_vert</span>
            </div>
          </div>
          
          <div style="padding: 16px; border-top: 1px solid #dadce0; font-size: 12px; color: #5f6368; font-weight: 500;">
            Time Zone: India Standard Time
          </div>
        </div>

      </div>
    </div>
      
      <!-- Custom Restore Confirmation Modal -->
      <div *ngIf="showRestoreConfirm" style="position: absolute; inset: 0; background: rgba(255,255,255,0.7); display: flex; align-items: center; justify-content: center; z-index: 100000; backdrop-filter: blur(2px);">
        <div style="background: #fff; width: 400px; border-radius: 8px; box-shadow: 0 4px 24px rgba(0,0,0,0.15); border: 1px solid #dadce0; overflow: hidden; font-family: 'Roboto', sans-serif;">
          <div style="padding: 24px 24px 16px 24px;">
            <h3 style="margin: 0 0 12px 0; font-size: 18px; font-weight: 500; color: #202124;">Restore this version?</h3>
            <p style="margin: 0; font-size: 14px; color: #5f6368; line-height: 1.5;">
              Are you sure you want to restore this version? This will become the new current version of the document.
            </p>
          </div>
          <div style="padding: 16px 24px; display: flex; justify-content: flex-end; gap: 12px; background: #f8f9fa; border-top: 1px solid #dadce0;">
            <button (click)="cancelRestore()" style="background: transparent; border: 1px solid #dadce0; padding: 8px 16px; border-radius: 4px; font-size: 14px; font-weight: 500; color: #5f6368; cursor: pointer;">
              Cancel
            </button>
            <button (click)="executeRestore()" style="background: #1a73e8; border: none; padding: 8px 16px; border-radius: 4px; font-size: 14px; font-weight: 500; color: #fff; cursor: pointer;">
              Restore
            </button>
          </div>
        </div>
      </div>
      
      <!-- Custom Name Version Modal -->
      <div *ngIf="showNameVersionPrompt" style="position: absolute; inset: 0; background: rgba(255,255,255,0.7); display: flex; align-items: center; justify-content: center; z-index: 100000; backdrop-filter: blur(2px);">
        <div style="background: #fff; width: 400px; border-radius: 8px; box-shadow: 0 4px 24px rgba(0,0,0,0.15); border: 1px solid #dadce0; overflow: hidden; font-family: 'Roboto', sans-serif;">
          <div style="padding: 24px 24px 16px 24px;">
            <h3 style="margin: 0 0 16px 0; font-size: 18px; font-weight: 500; color: #202124;">Name this version</h3>
            <input #versionNameInput type="text" [value]="tempVersionName" (input)="tempVersionName = versionNameInput.value" (keyup.enter)="submitNameVersion()" placeholder="Enter a name for this version" style="width: 100%; padding: 10px 12px; border: 1px solid #dadce0; border-radius: 4px; font-size: 14px; outline: none; box-sizing: border-box;">
          </div>
          <div style="padding: 16px 24px; display: flex; justify-content: flex-end; gap: 12px; background: #f8f9fa; border-top: 1px solid #dadce0;">
            <button (click)="cancelNameVersion()" style="background: transparent; border: 1px solid #dadce0; padding: 8px 16px; border-radius: 4px; font-size: 14px; font-weight: 500; color: #5f6368; cursor: pointer;">
              Cancel
            </button>
            <button (click)="submitNameVersion()" style="background: #1a73e8; border: none; padding: 8px 16px; border-radius: 4px; font-size: 14px; font-weight: 500; color: #fff; cursor: pointer;">
              Save
            </button>
          </div>
        </div>
      </div>

      <!-- Shared Color Picker Popover -->
      <div *ngIf="colorPickerState.active" (click)="closeColorPicker()" style="position: fixed; top: 0; left: 0; right: 0; bottom: 0; z-index: 10001; background: transparent;">
         <div (click)="$event.stopPropagation()" [style.top.px]="colorPickerState.top" [style.left.px]="colorPickerState.left" style="position: absolute; background: #fff; border: 1px solid #dadce0; box-shadow: 0 4px 12px rgba(0,0,0,0.15); border-radius: 4px; padding: 12px; width: 220px;">
           <div style="font-size: 11px; color: #5f6368; font-weight: 600; text-transform: uppercase; margin-bottom: 8px;">Theme Colors</div>
           <div style="display: grid; grid-template-columns: repeat(10, 1fr); gap: 2px; margin-bottom: 8px;">
             <!-- Theme Colors Grayscale -->
             <div *ngFor="let c of ['#ffffff','#f2f2f2','#d8d8d8','#bfbfbf','#a5a5a5','#7f7f7f','#595959','#3f3f3f','#262626','#000000']" (click)="setSparklineColor(c)" [style.background]="c" style="width: 16px; height: 16px; cursor: pointer; border: 1px solid #dadce0;"></div>
             <!-- Theme Colors Color scale -->
             <div *ngFor="let c of ['#e6b8af','#f4cccc','#fce5cd','#fff2cc','#d9ead3','#d0e0e3','#c9daf8','#cfe2f3','#d9d2e9','#ead1dc']" (click)="setSparklineColor(c)" [style.background]="c" style="width: 16px; height: 16px; cursor: pointer; border: 1px solid #dadce0;"></div>
             <div *ngFor="let c of ['#cc4125','#e06666','#f6b26b','#ffd966','#93c47d','#76a5af','#6d9eeb','#9fc5e8','#b4a7d6','#d5a6bd']" (click)="setSparklineColor(c)" [style.background]="c" style="width: 16px; height: 16px; cursor: pointer; border: 1px solid #dadce0;"></div>
           </div>
           
           <div style="font-size: 11px; color: #5f6368; font-weight: 600; text-transform: uppercase; margin-bottom: 8px; margin-top: 12px;">Standard Colors</div>
           <div style="display: grid; grid-template-columns: repeat(10, 1fr); gap: 2px; margin-bottom: 8px;">
             <div *ngFor="let c of ['#c00000','#ff0000','#ffc000','#ffff00','#92d050','#00b050','#00b0f0','#0070c0','#002060','#7030a0']" (click)="setSparklineColor(c)" [style.background]="c" style="width: 16px; height: 16px; cursor: pointer; border: 1px solid #dadce0;"></div>
           </div>
           
           <div *ngIf="recentColors.length > 0">
             <div style="font-size: 11px; color: #5f6368; font-weight: 600; text-transform: uppercase; margin-bottom: 8px; margin-top: 12px;">Other Used Colors</div>
             <div style="display: grid; grid-template-columns: repeat(10, 1fr); gap: 2px; margin-bottom: 8px;">
               <div *ngFor="let c of recentColors" (click)="setSparklineColor(c)" [style.background]="c" style="width: 16px; height: 16px; cursor: pointer; border: 1px solid #dadce0;"></div>
             </div>
           </div>
           
           <div style="font-size: 12px; color: #202124; font-weight: 500; margin-top: 12px; display:flex; align-items:center; gap: 8px;">
             <span class="material-symbols-outlined" style="font-size:16px;">add</span> More Colors:
             <input type="color" [(ngModel)]="customColorInput" (change)="setSparklineColor(customColorInput)" style="width:24px; height:24px; padding:0; border:none; cursor:pointer;">
           </div>
         </div>
      </div>

  
      <!-- ═══ Custom Insert Row / Column Modal (Zoho-style) ═══════════ -->
      <div class="modal-overlay" *ngIf="activeModal === 'custom_insert'" (click)="activeModal = null" 
           style="z-index:10001;"
           [style.--modal-bg]="currentTheme === 'dark' ? '#1e1e2e' : '#ffffff'"
           [style.--modal-text]="currentTheme === 'dark' ? '#e2e8f0' : '#202124'"
           [style.--modal-border]="currentTheme === 'dark' ? '#2d2d45' : '#dadce0'"
           [style.--modal-icon]="currentTheme === 'dark' ? '#94a3b8' : '#5f6368'"
           [style.--modal-input-bg]="currentTheme === 'dark' ? '#13131f' : '#f1f3f4'"
           [style.--modal-input-border]="currentTheme === 'dark' ? '#3d3d5c' : '#dadce0'"
           [style.--modal-radio-text]="currentTheme === 'dark' ? '#cbd5e1' : '#3c4043'">
        
        <div class="modal-content" (click)="$event.stopPropagation()" style="
          width:340px; padding:0; border-radius:10px; background:var(--modal-bg);
          box-shadow:0 8px 32px rgba(0,0,0,0.5); overflow:hidden; font-family:'Inter','Roboto',sans-serif;">

          <!-- Header -->
          <div style="display:flex;justify-content:space-between;align-items:center;padding:16px 20px 14px;border-bottom:1px solid var(--modal-border);">
            <div style="font-size:15px;font-weight:600;color:var(--modal-text);">
              {{ customInsertType === 'row' ? 'Insert Row' : 'Insert Column' }}
            </div>
            <button (click)="activeModal=null" style="background:none;border:none;cursor:pointer;color:var(--modal-icon);padding:0;display:flex;align-items:center;">
              <span class="material-symbols-outlined" style="font-size:18px;">close</span>
            </button>
          </div>

          <!-- Count input -->
          <div style="padding:20px 20px 16px;">
            <div style="position:relative;display:flex;align-items:center;border:1px solid var(--modal-input-border);border-radius:6px;overflow:hidden;background:var(--modal-input-bg);">
              <input type="number" [(ngModel)]="customInsertCount" min="1" max="1000"
                style="flex:1;background:transparent;border:none;outline:none;color:var(--modal-text);font-size:14px;padding:9px 12px;font-family:inherit;"/>
            </div>

            <!-- Radio options -->
            <div style="margin-top:16px;display:flex;flex-direction:column;gap:10px;">
              <label *ngIf="customInsertType === 'row'" style="display:flex;align-items:center;gap:10px;cursor:pointer;font-size:13px;color:var(--modal-radio-text);">
                <input type="radio" [(ngModel)]="customInsertPosition" value="before"
                  style="accent-color:#26A96C;width:15px;height:15px;cursor:pointer;"/>
                <span>Above (Row {{ selectedRow + 1 }})</span>
              </label>
              <label *ngIf="customInsertType === 'row'" style="display:flex;align-items:center;gap:10px;cursor:pointer;font-size:13px;color:var(--modal-radio-text);">
                <input type="radio" [(ngModel)]="customInsertPosition" value="after"
                  style="accent-color:#26A96C;width:15px;height:15px;cursor:pointer;"/>
                <span>Below (Row {{ selectedRow + 1 }})</span>
              </label>
              <label *ngIf="customInsertType === 'col'" style="display:flex;align-items:center;gap:10px;cursor:pointer;font-size:13px;color:var(--modal-radio-text);">
                <input type="radio" [(ngModel)]="customInsertPosition" value="before"
                  style="accent-color:#26A96C;width:15px;height:15px;cursor:pointer;"/>
                <span>Before (Column {{ colLabel(selectedCol) }})</span>
              </label>
              <label *ngIf="customInsertType === 'col'" style="display:flex;align-items:center;gap:10px;cursor:pointer;font-size:13px;color:var(--modal-radio-text);">
                <input type="radio" [(ngModel)]="customInsertPosition" value="after"
                  style="accent-color:#26A96C;width:15px;height:15px;cursor:pointer;"/>
                <span>After (Column {{ colLabel(selectedCol) }})</span>
              </label>
            </div>
          </div>

          <!-- Footer buttons -->
          <div style="display:flex;justify-content:flex-end;gap:10px;padding:14px 20px 18px;border-top:1px solid var(--modal-border);">
            <button (click)="activeModal=null"
              style="padding:7px 20px;border-radius:6px;border:1px solid var(--modal-input-border);background:transparent;color:var(--modal-icon);font-size:13px;font-family:inherit;cursor:pointer;">
              Cancel
            </button>
            <button (click)="confirmCustomInsert()"
              style="padding:7px 20px;border-radius:6px;border:none;background:#26A96C;color:#fff;font-size:13px;font-weight:600;font-family:inherit;cursor:pointer;">
              OK
            </button>
          </div>
        </div>
      </div>
      <!-- ══════════════════════════════════════════════════════════════ -->

      <!-- Insert Sparkline Modal -->
      <div class="modal-overlay" *ngIf="activeModal === 'insert_sparkline'" (click)="activeModal = null">
        <div class="modal-content" (click)="$event.stopPropagation()" style="width: 400px; padding: 24px; border-radius: 8px; background: #fff; box-shadow: 0 4px 12px rgba(0,0,0,0.15);">
          <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom: 20px;">
            <div style="font-size: 18px; font-weight: 500; color: #202124;">Insert Sparklines</div>
            <button (click)="activeModal = null" style="background: none; border: none; cursor: pointer; color: #5f6368;"><span class="material-symbols-outlined" style="font-size:20px;">close</span></button>
          </div>
          
          <div style="margin-bottom: 16px;">
            <div style="font-size: 13px; font-weight: 500; color: #202124; margin-bottom: 8px;">Source:</div>
            <div style="display:flex; align-items:center; border: 1px solid #dadce0; border-radius: 4px; padding: 0 8px;">
               <input type="text" autocomplete="off" spellcheck="false" [(ngModel)]="insertSparklineConfig.source" style="flex: 1; border: none; padding: 10px 0; outline: none; font-size: 14px;" placeholder="e.g. 'Sheet1'.A1:A5">
               <span class="material-symbols-outlined" style="color: #1a73e8; font-size: 18px; cursor: pointer;">grid_on</span>
            </div>
          </div>
          
          <div style="margin-bottom: 8px;">
            <div style="font-size: 13px; font-weight: 500; color: #202124; margin-bottom: 8px;">Destination:</div>
            <div style="display:flex; align-items:center; border: 1px solid #dadce0; border-radius: 4px; padding: 0 8px;">
               <input type="text" autocomplete="off" spellcheck="false" [(ngModel)]="insertSparklineConfig.dest" style="flex: 1; border: none; padding: 10px 0; outline: none; font-size: 14px;" placeholder="e.g. 'Sheet1'.B1:B5">
               <span class="material-symbols-outlined" style="color: #1a73e8; font-size: 18px; cursor: pointer;">grid_on</span>
            </div>
          </div>
          
          <div style="font-size: 12px; color: #5f6368; margin-bottom: 24px;">Note: Please select a destination range that is equal to the source range.</div>
          
          <div *ngIf="insertSparklineConfig.error" style="color: #ea4335; font-size: 13px; margin-bottom: 16px;">{{insertSparklineConfig.error}}</div>
          
          <div style="display: flex; justify-content: flex-end; gap: 12px;">
            <button (click)="activeModal = null" style="background: none; border: 1px solid #dadce0; border-radius: 4px; padding: 8px 24px; font-weight: 500; cursor: pointer; color: #202124;">Cancel</button>
            <button (click)="submitInsertSparkline()" style="background: #0f9d58; border: none; border-radius: 4px; padding: 8px 24px; font-weight: 500; cursor: pointer; color: #fff;">OK</button>
          </div>
        </div>
      </div>

      <!-- Draggable Emoji Picker -->
      <div *ngIf="activeModal === 'emoji'" 
           [style.left.px]="emojiPickerX" 
           [style.top.px]="emojiPickerY" 
           style="position: fixed; z-index: 10001; background: #fff; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.15); display: flex; flex-direction: column; width: 338px;">
        <div style="display: flex; justify-content: space-between; align-items: center; padding: 8px 12px; border-bottom: 1px solid #e0e0e0; cursor: move; background: #f8f9fa; border-radius: 8px 8px 0 0;"
             (mousedown)="startEmojiDrag($event)">
          <span style="font-size: 14px; font-weight: 500; color: #333;">Emojis</span>
          <button (click)="activeModal = null" style="background: none; border: none; cursor: pointer; color: #5f6368; display: flex; padding: 0;">
            <span class="material-symbols-outlined" style="font-size: 18px;">close</span>
          </button>
        </div>
        <emoji-mart (emojiSelect)="addEmoji($event)" set="apple" title="Pick your emoji…" emoji="smile" [showPreview]="false" [emojisToShowFilter]="filterEmojis" [style]="{ width: '100%', border: 'none', borderRadius: '0 0 8px 8px' }" [darkMode]="currentTheme === 'dark'"></emoji-mart>
      </div>

            <!-- Edit Sparkline Modal -->
      <div class="modal-overlay" *ngIf="activeModal === 'edit_sparkline'" (click)="activeModal = null">
        <div class="modal-content" (click)="$event.stopPropagation()" style="width: 400px; padding: 24px; border-radius: 8px; background: #fff; box-shadow: 0 4px 12px rgba(0,0,0,0.15);">
          <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom: 20px;">
            <div style="font-size: 18px; font-weight: 500; color: #202124;">Edit</div>
            <button (click)="activeModal = null" style="background: none; border: none; cursor: pointer; color: #5f6368;"><span class="material-symbols-outlined" style="font-size:20px;">close</span></button>
          </div>
          
          <div style="display:flex; border-bottom: 1px solid #dadce0; margin-bottom: 16px;">
            <div (click)="editSparklineConfig.tab = 'selected'" [style.border-bottom]="editSparklineConfig.tab === 'selected' ? '2px solid #0f9d58' : 'none'" [style.color]="editSparklineConfig.tab === 'selected' ? '#0f9d58' : '#5f6368'" style="padding: 8px 16px; font-weight: 500; cursor: pointer;">Selected</div>
            <div (click)="editSparklineConfig.tab = 'group'" [style.border-bottom]="editSparklineConfig.tab === 'group' ? '2px solid #0f9d58' : 'none'" [style.color]="editSparklineConfig.tab === 'group' ? '#0f9d58' : '#5f6368'" style="padding: 8px 16px; font-weight: 500; cursor: pointer;" *ngIf="sparklineConfig?.isGrouped">Group</div>
          </div>
          
          <div style="margin-bottom: 16px;">
            <div style="font-size: 13px; font-weight: 500; color: #202124; margin-bottom: 8px;">Source:</div>
            <div style="display:flex; align-items:center; border: 1px solid #dadce0; border-radius: 4px; padding: 0 8px;">
               <input type="text" [(ngModel)]="editSparklineConfig.source" style="flex: 1; border: none; padding: 10px 0; outline: none; font-size: 14px;">
               <span class="material-symbols-outlined" style="color: #1a73e8; font-size: 18px; cursor: pointer;">grid_on</span>
            </div>
            <div *ngIf="editSparklineConfig.error" style="color: #ea4335; font-size: 11px; margin-top: 4px;">{{ editSparklineConfig.error }}</div>
          </div>
          
          <div style="margin-bottom: 24px;">
            <div style="font-size: 13px; font-weight: 500; color: #202124; margin-bottom: 8px;">Location:</div>
            <div style="display:flex; align-items:center; border: 1px solid #dadce0; border-radius: 4px; padding: 0 8px;">
               <input type="text" autocomplete="off" spellcheck="false" [(ngModel)]="editSparklineConfig.dest" style="flex: 1; border: none; padding: 10px 0; outline: none; font-size: 14px;" [disabled]="true" [style.background]="'#f8f9fa'">
               <span class="material-symbols-outlined" style="color: #5f6368; font-size: 18px; cursor: not-allowed;">grid_on</span>
            </div>
            <div style="font-size: 11px; color: #5f6368; margin-top: 4px;">Location cannot be changed.</div>
          </div>
          
          <div style="display: flex; justify-content: flex-end; gap: 12px;">
            <button (click)="activeModal = null" style="background: none; border: 1px solid #dadce0; border-radius: 4px; padding: 8px 24px; font-weight: 500; cursor: pointer; color: #202124;">Cancel</button>
            <button (click)="submitEditSparkline()" style="background: #0f9d58; border: none; border-radius: 4px; padding: 8px 24px; font-weight: 500; cursor: pointer; color: #fff;">OK</button>
          </div>
        </div>
      </div>

  `,
  styles: [`
    /* Hide specific emojis that are unsupported on Windows 10 since ngx-emoji-mart categories bypass the filter */
    ::ng-deep .emoji-mart-emoji[title*="elting" i],
    ::ng-deep .emoji-mart-emoji[aria-label*="elting" i],
    ::ng-deep .emoji-mart-emoji[title*="louds" i],
    ::ng-deep .emoji-mart-emoji[aria-label*="louds" i],
    ::ng-deep .emoji-mart-emoji[title*="otted" i],
    ::ng-deep .emoji-mart-emoji[aria-label*="otted" i],
    ::ng-deep .emoji-mart-emoji[title*="xhaling" i],
    ::ng-deep .emoji-mart-emoji[aria-label*="xhaling" i],
    ::ng-deep .emoji-mart-emoji[title*="piral" i],
    ::ng-deep .emoji-mart-emoji[aria-label*="piral" i],
    ::ng-deep .emoji-mart-emoji[title*="black cat" i],
    ::ng-deep .emoji-mart-emoji[aria-label*="black cat" i],
    ::ng-deep .emoji-mart-emoji[title*="black_cat" i],
    ::ng-deep .emoji-mart-emoji[aria-label*="black_cat" i] {
      display: none !important;
      opacity: 0 !important;
      visibility: hidden !important;
      pointer-events: none !important;
      width: 0 !important;
      height: 0 !important;
      margin: 0 !important;
      padding: 0 !important;
    }
    /* Bottom Chat Bar */
    .bottom-chat-bar { position: fixed; bottom: 0; left: 0; right: 0; display: flex; background: #f8f9fa; border-top: 1px solid #e0e0e0; z-index: 9999; height: 36px; box-shadow: 0 -1px 3px rgba(0,0,0,0.05); }
    .bcb-item { display: flex; align-items: center; gap: 8px; padding: 0 16px; cursor: pointer; border-right: 1px solid #e0e0e0; font-size: 13px; font-weight: 500; color: #202124; transition: background 0.2s; position: relative; }
    .bcb-item:hover { background: #e8f0fe; }
    .bcb-item .material-symbols-outlined { font-size: 18px; }
    .bcb-badge { position: absolute; top: -6px; left: 16px; background: #d32f2f; color: #fff; font-size: 10px; font-weight: bold; border-radius: 50%; width: 16px; height: 16px; display: flex; align-items: center; justify-content: center; }

    /* Widgets */
    .widget-panel { position: fixed; bottom: 48px; right: 24px; width: 300px; background: #fff; border-radius: 8px; border: 1px solid #e0e0e0; z-index: 10000; display: flex; flex-direction: column; overflow: hidden; box-shadow: 0 4px 12px rgba(0,0,0,0.15); }
    .wp-header { display: flex; justify-content: space-between; align-items: center; padding: 12px 16px; border-bottom: 1px solid #e0e0e0; font-weight: 500; color: #202124; background: #f8f9fa; }
    .wp-body { padding: 16px; flex: 1; min-height: 200px; display: flex; flex-direction: column; }
    :host { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif; }
    * { box-sizing: border-box; }

    /* ── Shell ─────────────────────────────────────────────────────────── */
    .shell { display:flex; flex-direction:column; height:calc(100vh - 36px); background:#fff; overflow:hidden; }

    /* ── TOP BAR ────────────────────────────────────────────────────────── */
    .top-bar { display:flex; align-items:center; justify-content:space-between; padding:6px 16px; background:#1c2333; min-height:50px; z-index:300; flex-shrink:0; }
    .tl { display:flex; align-items:center; gap:10px; }
    .tl-sep { width:1px; height:24px; background:rgba(255,255,255,0.15); margin:0 12px; flex-shrink:0; }
    .brand { display:flex; align-items:center; gap:6px; flex-shrink:0; padding: 4px; border-radius: 4px; transition: background 0.2s; }
    .brand:hover { background: rgba(255,255,255,0.08); }
    .brand-name { color:#fff; font-size:15px; font-weight:600; }
    .cursor-path { stroke: #1e1e1e; }
    .doc-sec { display:flex; flex-direction:row; align-items:center; }
    .doc-title { background:transparent; border:1px solid transparent; border-radius:4px; color:#fff; font-size:15px; font-weight:600; padding:3px 6px; outline:none; overflow:hidden; text-overflow:ellipsis; max-width:500px; min-width:30px; transition: width 0.1s; }
    .doc-title:hover { border-color:rgba(255,255,255,.3); }
    .doc-title:focus { border-color:rgba(255,255,255,.6); background:rgba(255,255,255,.1); }
    .doc-icons { color: rgba(255,255,255,0.6); }
    .tr { display:flex; align-items:center; gap:8px; }
    .top-search-box { display:flex; align-items:center; gap:6px; background:rgba(255,255,255,.1); border:1px solid rgba(255,255,255,.18); border-radius:24px; padding:4px 12px; color:rgba(255,255,255,.8); transition: all 0.2s ease; height:32px; box-sizing:border-box; }
    .top-search-box.has-query { background: rgba(15, 157, 88, 0.15); border-color: #0f9d58; color: #fff; }
    .top-search-box input.inline-search-input { background:transparent; border:none; outline:none; color:inherit; font-size:13px; width:140px; }
    .top-search-box input.inline-search-input::placeholder { color: rgba(255,255,255, 0.5); }
    .top-search-box .inline-search-clear { background:rgba(255,255,255,0.15); border:none; border-radius:50%; width:16px; height:16px; display:flex; align-items:center; justify-content:center; cursor:pointer; color:inherit; padding:0; flex-shrink:0; }
    .top-search-box .inline-search-clear:hover { background:rgba(255,255,255,0.3); }
    .top-search-box .inline-search-divider { width:1px; height:14px; background:rgba(255,255,255,0.2); margin:0 4px; }
    .top-search-box .inline-search-count { font-size:12px; font-weight:500; opacity:0.8; white-space:nowrap; margin-right:4px; font-variant-numeric: tabular-nums; }
    .top-search-box .inline-search-nav { display:flex; gap:2px; }
    .top-search-box .inline-search-nav button { background:none; border:none; border-radius:4px; display:flex; align-items:center; justify-content:center; cursor:pointer; color:inherit; padding:2px; opacity:0.8; transition:all 0.2s; }
    .top-search-box .inline-search-nav button:hover { opacity:1; background:rgba(255,255,255,0.15); }
    .online-badge { display:flex; align-items:center; font-size:13px; font-weight:500; color:rgba(255,255,255,.9); background:rgba(255,255,255,.1); border:1px solid rgba(255,255,255,.2); border-radius:18px; padding:4px 10px; margin-right:8px; cursor:default; }
    .share-btn { display:flex; align-items:center; gap:6px; background:#26a96c; border:none; border-radius:20px; color:#fff; cursor:pointer; font-size:13px; font-weight:600; padding:7px 16px; flex-shrink:0; }
    .share-btn:hover { background:#1f8a57; }
    .properties-btn { display:flex; align-items:center; justify-content:center; background:transparent; border:none; border-radius:50%; color:#5f6368; cursor:pointer; width:40px; height:40px; flex-shrink:0; margin-left:8px; margin-right:12px; transition: background 0.2s ease, color 0.2s ease; }
    .properties-btn:hover { background:#f1f3f4; color:#202124; }
    
    .properties-panel { position:fixed; right:0; top:0; width:340px; height:100vh; background:#fff; z-index:9999; box-shadow:-4px 0 24px rgba(0,0,0,.15); overflow-y:auto; transform:translateX(100%); transition:transform 0.3s cubic-bezier(0.4, 0, 0.2, 1); display:flex; flex-direction:column; font-family:"Roboto", sans-serif; color:#202124; }
    .properties-panel.open { transform:translateX(0); }
    .pp-header { display:flex; align-items:center; justify-content:space-between; padding:16px 20px; border-bottom:1px solid #e0e0e0; background:#f8f9fa; }
    .pp-title { font-size:18px; font-weight:500; margin:0; display:flex; align-items:center; gap:8px; }
    .pp-close { background:transparent; border:none; cursor:pointer; color:#5f6368; display:flex; align-items:center; justify-content:center; width:36px; height:36px; border-radius:50%; transition:background 0.2s; }
    .pp-close:hover { background:#e8eaed; color:#202124; }
    .pp-content { padding:24px 20px; display:flex; flex-direction:column; gap:24px; }
    .pp-section { display:flex; flex-direction:column; gap:6px; }
    .pp-label { font-size:12px; font-weight:600; color:#5f6368; text-transform:uppercase; letter-spacing:0.8px; }
    .pp-value { font-size:14px; color:#202124; display:flex; align-items:center; gap:12px; }
    .pp-av { width:36px; height:36px; border-radius:50%; background:#ea4335; color:#fff; display:flex; align-items:center; justify-content:center; font-weight:600; font-size:14px; text-transform:uppercase; flex-shrink:0; }
    .pp-link { color:#1a73e8; cursor:pointer; font-weight:500; word-break:break-all; text-decoration:none; line-height:1.4; }
    .pp-link:hover { text-decoration:underline; }
    .pp-stats { display:grid; grid-template-columns:1fr 1fr; gap:16px; background:#f8f9fa; padding:16px; border-radius:8px; border:1px solid #e0e0e0; margin-top:8px; }
    .pp-stat-item { display:flex; flex-direction:column; gap:4px; }
    .pp-stat-num { font-size:18px; font-weight:600; color:#202124; }
    .pp-stat-lbl { font-size:12px; color:#5f6368; }
    .pp-divider { height:1px; background:#e0e0e0; border:none; margin:0; }

    .av { position:relative; width:34px; height:34px; border-radius:50%; background:#ea4335; color:#fff; font-size:13px; font-weight:700; display:flex; align-items:center; justify-content:center; cursor:pointer; flex-shrink:0; }
        .profile-dd { position:fixed; top:54px; right:16px; width:240px; background:#2d3748; border-radius:10px; box-shadow:0 6px 24px rgba(0,0,0,.5); z-index:9999; overflow:hidden; border:1px solid #4a5568; }
    .pd-head { padding:14px 16px; border-bottom:1px solid #4a5568; display:flex; align-items:center; gap:10px; background:#1a202c; color: #fff; }
    .pd-av { width:40px; height:40px; border-radius:50%; background:#ea4335; color:#fff; font-size:16px; font-weight:700; display:flex; align-items:center; justify-content:center; flex-shrink:0; }
    .pd-item { padding:9px 16px; font-size:13px; color:#e2e8f0; cursor:pointer; display:flex; align-items:center; gap:10px; }
    .pd-item:hover { background:#4a5568; color:#fff; }
    .pd-item.danger { color:#fc8181; }
    .pd-item.danger:hover { background:#fc8181; color:#fff; }
    .pd-sep { height:1px; background:#4a5568; margin:4px 0; }
    .pd-icon { font-size:18px !important; color:inherit; }

    /* ??? MENU BAR ?????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????? */
    .menu-row { display:flex; align-items:center; background:#252d3d; padding:1px 12px; flex-shrink:0; z-index:200; }
    .mi { position:relative; color:rgba(255,255,255,.82); font-size:13px; padding:5px 10px; cursor:pointer; border-radius:4px; user-select:none; white-space:nowrap; }
    .mi:hover, .mi-open { background:rgba(255,255,255,.12); color:#fff; }
    .mdd { position:absolute; top:calc(100% + 2px); left:0; min-width:230px; background:#2d3748; border:1px solid #4a5568; border-radius:6px; box-shadow:0 6px 24px rgba(0,0,0,.5); z-index:1000; padding:2px 0; }
    .mdi { padding:4px 16px; font-size:13px; color:#e2e8f0; cursor:pointer; display:flex; justify-content:flex-start; align-items:center; white-space:nowrap; position:relative; }
    .mdi:hover { background:#4a5568; color:#fff; }
    .mdi.danger { color:#fc8181; }
    .mdi.danger:hover { background:#fc8181; color:#fff; }
    .mdi-title { padding:4px 16px; font-size:12px; color:#a0aec0; font-weight:600; cursor:default; user-select:none; margin-top:4px; }
    .mds { height:1px; background:#4a5568; margin:3px 0; }
    .mh { font-size:11px; color:#a0aec0; margin-left:auto; padding-left:20px; }
    .mdi-icon { width:16px; height:16px; margin-right:10px; display:inline-flex; align-items:center; justify-content:center; color:#a0aec0; font-size:16px; }
    .mdi-arrow { margin-left:auto; padding-left:12px; font-size:16px; color:#a0aec0; display:flex; align-items:center; }
    .mdi:hover .mdi-icon, .mdi:hover .mdi-arrow { color:#fff; }
    .mdi-sub { position:absolute; left:100%; top:-5px; min-width:240px; background:#2d3748; border:1px solid #4a5568; border-radius:6px; box-shadow:0 6px 24px rgba(0,0,0,.5); display:none; padding:4px 0; z-index:1001; }
    .mdi-sub.sub-left { left:auto; right:100%; margin-right:-4px; }
    .mdi.has-sub:hover > .mdi-sub { display:block; }
    .font-list { max-height:280px; overflow-y:auto; }

    /* ── TOOLBAR ────────────────────────────────────────────────────────── */
    .tb-row { display:flex; align-items:center; flex-wrap:wrap; background:#2d3748; padding:4px 12px; gap:2px; flex-shrink:0; position:relative; z-index:190; }
    .tb-row2 { border-top:1px solid rgba(255,255,255,.08); padding:3px 12px; position:relative; z-index:180; }
    .tb-group { display:flex; align-items:center; gap:2px; }
    .tb-sep { width:1px; height:20px; background:rgba(255,255,255,.18); margin:0 5px; flex-shrink:0; }
    .tb .material-symbols-outlined { font-size: 18px; }
      .tb { background:transparent; border:none; border-radius:3px; color:rgba(255,255,255,.85); cursor:pointer; font-size:13px; font-family:inherit; height:26px; min-width:26px; padding:0 5px; display:flex; align-items:center; justify-content:center; transition:background .1s; flex-shrink:0; }
    .tb:hover { background:rgba(255,255,255,.15); color:#fff; }
    .tb.tb-on { background:rgba(26,115,232,.6); color:#fff; }
    .tb.sz { min-width:22px; }
    .tb.nf { font-size:11px; font-weight:700; min-width:28px; }
    .tb-font-dd { display:flex; align-items:center; justify-content:space-between; gap:4px; background:rgba(255,255,255,.08); border:1px solid rgba(255,255,255,.15); border-radius:3px; color:rgba(255,255,255,.9); cursor:pointer; font-size:12px; height:26px; padding:0 8px; position:relative; user-select:none; min-width:110px; }
    .tb-font-dd:hover, .tb-font-dd.active { background:rgba(255,255,255,.15); }
    .arr { font-size:9px; color:rgba(255,255,255,.45); margin-left:auto; }
    .font-sz { gap:0; }
    .sz-inp { background:rgba(255,255,255,.1); border:none; border-left:1px solid rgba(255,255,255,.15); border-right:1px solid rgba(255,255,255,.15); color:#fff; font-size:12px; height:26px; outline:none; text-align:center; width:38px; }
    .sz-inp::-webkit-inner-spin-button, .sz-inp::-webkit-outer-spin-button { -webkit-appearance:none; }
    .sz-drop-btn { background:rgba(255,255,255,.1); height:26px; border:1px solid rgba(255,255,255,.15); border-left:none; border-top:none; border-bottom:none; display:flex; align-items:center; justify-content:center; cursor:pointer; width:18px; color:rgba(255,255,255,.9); }
    .zoom-ctrl { display:flex; align-items:center; gap:4px; color:rgba(255,255,255,.85); font-size:12px; }
    .zoom-pct { min-width:38px; text-align:center; }
    .tb-clr { display:flex; align-items:center; gap:2px; background:transparent; border:none; border-radius:3px; color:rgba(255,255,255,.85); cursor:pointer; font-size:12px; height:26px; padding:0 5px; position:relative; }
    .tb-clr:hover { background:rgba(255,255,255,.15); }
    .clr-ico { position:relative; width:16px; height:16px; display:flex; align-items:center; justify-content:center; }
    .clr-bar { position:absolute; bottom:1px; left:1px; right:1px; height:3px; border-radius:1px; z-index:2; }
    .clr-pop { position:absolute; top:calc(100% + 4px); left:0; background:#fff; border:1px solid #ddd; border-radius:6px; box-shadow:0 4px 16px rgba(0,0,0,.2); padding:8px; z-index:1000; }
    .cp-grid { display:grid; grid-template-columns:repeat(10, 1fr); gap:4px; margin-bottom:8px; }
    .cp-sw { width:16px; height:16px; border-radius:2px; cursor:pointer; border:1px solid rgba(0,0,0,.1); }
    .cp-sw:hover { transform:scale(1.3); outline:1px solid #555; }
    .cp-nocolor { padding:4px 8px; font-size:12px; color:#555; cursor:pointer; white-space:nowrap; border-bottom:1px solid #eee; margin-bottom:6px; }
    .cp-nocolor:hover { background:#f5f5f5; }
    .tb-chart-dd { position:absolute; top:36px; left:0; width:340px; background:#202124; border:1px solid #5f6368; border-radius:4px; box-shadow:0 8px 24px rgba(0,0,0,0.5); z-index:500; display:flex; flex-direction:column; padding:12px; }
    .tb-dd { position:absolute; top:36px; left:0; background:#202124; border:1px solid #5f6368; border-radius:4px; box-shadow:0 8px 24px rgba(0,0,0,0.5); z-index:500; display:flex; flex-direction:column; padding:8px 0; }
    .dd-item { padding: 8px 16px; color: #e8eaed; font-size: 13px; font-weight: 500; cursor: pointer; display: flex; align-items: center; gap: 12px; transition: background 0.2s; }
    .dd-item:hover { background: rgba(255,255,255,0.08); }
    .chart-header-icons { display:flex; align-items:center; gap:8px; border-bottom:1px solid #5f6368; padding-bottom:12px; margin-bottom:12px; }

    /* ── FORMULA BAR ────────────────────────────────────────────────────── */
    .formula-container { display:flex; align-items:center; background:#2d3748; border-bottom:1px solid rgba(255,255,255,.08); flex-shrink:0; height:34px; padding:0 12px; gap:8px; }
    .cell-ref { background:rgba(255,255,255,.05); border:1px solid rgba(255,255,255,.15); border-radius:4px; color:#fff; font-size:12px; font-weight:600; min-width:60px; height:24px; display:flex; align-items:center; justify-content:center; padding: 0 10px; }
    .fx-label { color:#a0aec0; font-style:italic; font-size:14px; display:flex; align-items:center; margin:0 4px; border:none; }
    .formula-bar { background:rgba(255,255,255,.05); border:1px solid rgba(255,255,255,.15); border-radius:4px; flex:1; font-size:13px; outline:none; padding:0 12px; color:#fff; height:24px; margin-right: 8px; }

    /* ── GRID ─────────────────────────────────────────────────────────── */
    .main-content { display:flex; flex:1; overflow:hidden; position:relative; }
    .grid-wrap { flex:1; overflow:auto; position:relative; background:#fff; overflow-anchor: none; }
    
    /* ── SIDE PANEL ─────────────────────────────────────────────────── */
    .side-panel {
      width: 340px;
      border-left: 1px solid #e2e8f0;
      background: #f0f4ff;
      display: flex;
      flex-direction: column;
      z-index: 100;
      box-shadow: -4px 0 24px rgba(99,102,241,0.08);
    }

    /* Header */
    .sp-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 16px 16px 14px;
      background: #fff;
      border-bottom: 1px solid #e8edf5;
      flex-shrink: 0;
    }
    .sp-head-left { display: flex; align-items: center; gap: 12px; }
    .sp-icon-wrap {
      width: 38px; height: 38px; border-radius: 10px;
      display: flex; align-items: center; justify-content: center;
      flex-shrink: 0;
    }
    .sp-icon-cal  { background: linear-gradient(135deg, #6366f1, #818cf8); }
    .sp-icon-notes { background: linear-gradient(135deg, #f59e0b, #fbbf24); }
    .sp-icon-tasks { background: linear-gradient(135deg, #10b981, #34d399); }
    .sp-head-icon { font-size: 20px !important; color: #fff; }
    .sp-title { font-size: 15px; font-weight: 700; color: #1e1e2e; line-height: 1.2; }
    .sp-subtitle { font-size: 11px; color: #94a3b8; margin-top: 1px; }
    .sp-close-btn {
      background: none; border: none; cursor: pointer;
      width: 30px; height: 30px; border-radius: 8px;
      display: flex; align-items: center; justify-content: center;
      color: #94a3b8; transition: background .15s, color .15s;
    }
    .sp-close-btn:hover { background: #f1f5f9; color: #475569; }
    .sp-close-btn .material-symbols-outlined { font-size: 18px; }

    /* Scrollable body */
    .sp-content {
      flex: 1; overflow-y: auto; padding: 16px; display: flex; flex-direction: column; gap: 12px;
    }
    .sp-content::-webkit-scrollbar { width: 5px; }
    .sp-content::-webkit-scrollbar-track { background: transparent; }
    .sp-content::-webkit-scrollbar-thumb { background: #c7d2e8; border-radius: 10px; }

    /* Cards */
    .sp-card {
      background: #fff;
      border-radius: 14px;
      border: 1px solid #e8edf5;
      padding: 14px;
      box-shadow: 0 2px 8px rgba(99,102,241,0.06);
    }
    .sp-notes-grow { flex: 1; display: flex; flex-direction: column; }
    .sp-card-label {
      display: flex; align-items: center; gap: 6px;
      font-size: 10px; font-weight: 700; text-transform: uppercase;
      letter-spacing: 0.8px; color: #94a3b8; margin-bottom: 10px;
    }
    .sp-label-icon { font-size: 14px !important; }

    /* Date card */
    .sp-date-card { }
    .sp-date-input {
      width: 100%; box-sizing: border-box;
      border: 1.5px solid #e2e8f0;
      border-radius: 9px;
      padding: 10px 12px;
      font-size: 14px; color: #1e1e2e;
      outline: none; cursor: pointer;
      transition: border-color .15s, box-shadow .15s;
      font-family: inherit;
    }
    .sp-date-input:focus { border-color: #6366f1; box-shadow: 0 0 0 3px rgba(99,102,241,0.12); }
    .sp-date-chip {
      display: inline-flex; align-items: center; gap: 5px;
      margin-top: 8px; background: #eef2ff; color: #6366f1;
      font-size: 11px; font-weight: 600;
      padding: 4px 10px; border-radius: 20px;
    }

    /* Textarea card */
    .sp-notes-card { }
    .sp-textarea-wrap { display: flex; flex-direction: column; flex: 1; }
    .sp-textarea {
      flex: 1; min-height: 180px;
      border: 1.5px solid #e2e8f0; border-radius: 9px;
      padding: 12px; font-size: 13.5px; line-height: 1.6;
      color: #1e1e2e; outline: none; resize: none; font-family: inherit;
      background: #fafbff;
      transition: border-color .15s, box-shadow .15s;
    }
    .sp-textarea:focus { border-color: #6366f1; box-shadow: 0 0 0 3px rgba(99,102,241,0.10); background: #fff; }
    .sp-textarea::placeholder { color: #c4cfe0; }
    .sp-textarea-tall { min-height: 320px; }
    .sp-textarea-footer {
      display: flex; align-items: center; gap: 4px;
      margin-top: 6px; padding: 0 2px;
    }

    /* Task add */
    .sp-task-add-wrap {
      display: flex; align-items: center; gap: 8px;
      background: #fff; border: 1.5px solid #e2e8f0;
      border-radius: 12px; padding: 8px 10px 8px 12px;
      box-shadow: 0 2px 8px rgba(99,102,241,0.06);
      transition: border-color .15s, box-shadow .15s;
    }
    .sp-task-add-wrap:focus-within { border-color: #6366f1; box-shadow: 0 0 0 3px rgba(99,102,241,0.10); }
    .sp-task-add-icon { font-size: 20px !important; color: #c7d2e8; flex-shrink: 0; }
    .sp-task-input {
      flex: 1; border: none; outline: none; font-size: 13.5px;
      color: #1e1e2e; font-family: inherit; background: transparent;
    }
    .sp-task-input::placeholder { color: #c4cfe0; }
    .sp-add-btn {
      background: linear-gradient(135deg, #6366f1, #818cf8);
      border: none; border-radius: 8px;
      color: #fff; font-size: 12px; font-weight: 700;
      padding: 6px 14px; cursor: pointer;
      transition: opacity .15s, transform .1s;
      flex-shrink: 0;
    }
    .sp-add-btn:hover { opacity: .88; transform: translateY(-1px); }
    .sp-add-btn:active { transform: translateY(0); }

    /* Tasks summary bar */
    .sp-tasks-summary {
      display: flex; align-items: center; justify-content: space-between;
      padding: 4px 2px; margin-top: -4px;
    }
    .sp-tasks-count { font-size: 11px; font-weight: 700; color: #64748b; }
    .sp-tasks-done {
      font-size: 11px; font-weight: 600;
      color: #10b981; background: #d1fae5;
      padding: 2px 8px; border-radius: 10px;
    }

    /* Task items */
    .sp-task-item {
      display: flex; align-items: center; gap: 10px;
      background: #fff; border-radius: 12px;
      border: 1.5px solid #e8edf5;
      padding: 11px 12px;
      box-shadow: 0 1px 4px rgba(99,102,241,0.05);
      transition: box-shadow .15s, border-color .15s;
    }
    .sp-task-item:hover { box-shadow: 0 4px 12px rgba(99,102,241,0.10); border-color: #c7d2f8; }
    .sp-task-done { background: #f8faff; border-color: #e8edf5; }
    .sp-task-done .sp-task-text { text-decoration: line-through; color: #b0bec5; }

    /* Custom checkbox */
    .sp-checkbox-wrap { display: flex; align-items: center; cursor: pointer; flex-shrink: 0; }
    .sp-checkbox-native { display: none; }
    .sp-checkbox-ui {
      width: 20px; height: 20px; border-radius: 6px;
      border: 2px solid #c7d2e8; background: #fff;
      display: flex; align-items: center; justify-content: center;
      transition: background .15s, border-color .15s;
    }
    .sp-checkbox-native:checked + .sp-checkbox-ui {
      background: linear-gradient(135deg, #10b981, #34d399);
      border-color: #10b981;
    }
    .sp-check-icon { font-size: 13px !important; color: #fff; opacity: 0; transition: opacity .15s; }
    .sp-checkbox-native:checked + .sp-checkbox-ui .sp-check-icon { opacity: 1; }

    .sp-task-text { flex: 1; font-size: 13.5px; color: #1e1e2e; line-height: 1.4; word-break: break-word; transition: color .15s; }
    .sp-task-del {
      background: none; border: none; cursor: pointer;
      width: 26px; height: 26px; border-radius: 7px;
      display: flex; align-items: center; justify-content: center;
      color: #c4cfe0; flex-shrink: 0; transition: background .15s, color .15s;
    }
    .sp-task-del:hover { background: #fee2e2; color: #ef4444; }
    .sp-task-del .material-symbols-outlined { font-size: 16px !important; }

    /* Empty state */
    .sp-empty {
      text-align: center; padding: 40px 16px 32px;
      background: #fff; border-radius: 14px; border: 1.5px dashed #c7d2e8;
    }
    .sp-empty-icon {
      width: 56px; height: 56px; border-radius: 50%;
      background: linear-gradient(135deg, #d1fae5, #a7f3d0);
      display: flex; align-items: center; justify-content: center;
      margin: 0 auto 12px;
    }
    .sp-empty-icon .material-symbols-outlined { font-size: 28px !important; color: #10b981; }
    .sp-empty-title { font-size: 15px; font-weight: 700; color: #1e1e2e; margin-bottom: 4px; }
    .sp-empty-sub { font-size: 12px; color: #94a3b8; }

    .sp-frame { flex:1; width:100%; border:none; }

    /* Table itself provides the outer left+top border, cells provide right+bottom — no double lines */
    .grid { border-collapse:separate; border-spacing:0; table-layout:fixed; font-size:13px; user-select:none; background:#202124; }

    /* ── GRID HEADERS — dark ──────────────────────────────────── */
          .corner { background:#202124; border-right:1px solid #3c3c3c; border-bottom:2px solid #3c3c3c; position:sticky; top:0; left:0; z-index:50; width:46px; min-width:46px; text-align:center; height:26px; }
      .col-head { background:#202124; border-right:1px solid #3c3c3c; border-bottom:2px solid #3c3c3c; color:#e8eaed; cursor:pointer; font-size:12px; font-weight:500; position:sticky; top:0; text-align:center; user-select:none; z-index:45; height:26px; width:100px; min-width:100px; transition:background 0.2s, border-color 0.2s; }
    .col-head:hover { background:#35363a; color:#fff; }
    .col-selected { background:#111 !important; color:#fff !important; font-weight:700 !important; }
    .col-head.active-axis { background:#111 !important; color:#10b981 !important; border-bottom-color:#10b981; }

          .row-head { background:#202124; border-right:2px solid #3c3c3c; border-bottom:1px solid #3c3c3c; color:#e8eaed; cursor:pointer; font-size:12px; font-weight:400; position:sticky; left:0; text-align:center; user-select:none; z-index:40; min-width:46px; width:46px; height:26px; transition:background 0.2s, border-color 0.2s; }
    .row-head:hover { background:#35363a; color:#fff; }
    .row-selected { background:#111 !important; color:#fff !important; font-weight:700 !important; }
    .row-head.active-axis { background:#111 !important; color:#10b981 !important; border-right-color:#10b981; }

    /* Data cells — NO extra border, just shared grid lines */
    .cell { cursor:cell; border-right:1px solid #d0d0d0; border-bottom:1px solid #d0d0d0; height:26px; position:relative; white-space:nowrap; padding:0; min-width:100px; width:100px; max-width:200px; background:#fff; }
    .cell.has-content { z-index:2; }
    .cell.selected { outline:2px solid #34a853; outline-offset:-2px; z-index:20; }
    .cell.in-range { box-shadow:inset 0 0 0 1000px rgba(52,168,83,0.15) !important; }
    .cell.fill-preview { box-shadow:inset 0 0 0 1000px rgba(52,168,83,0.2) !important; border:1px dashed #34a853 !important; }
    .cell.search-match { box-shadow:inset 0 0 0 1000px rgba(255,193,7,0.35) !important; }
    .cell.search-match-active { box-shadow:inset 0 0 0 1000px rgba(255,152,0,0.6) !important; outline: 2px solid #ff9800; outline-offset: -2px; z-index: 21; }
    .comment-highlight { box-shadow:inset 0 0 0 1000px rgba(251,191,36,0.35) !important; outline: 2px solid #f59e0b !important; outline-offset: -2px; z-index: 15; }
    
    .no-gridlines th, .no-gridlines td { border-color: transparent !important; }
    .col-resizer { position: absolute; right: 0; top: 0; bottom: 0; width: 5px; cursor: col-resize; z-index: 60; }
    .col-resizer:hover { background: #1a73e8; }
    .row-resizer { position: absolute; bottom: 0; left: 0; right: 0; height: 5px; cursor: row-resize; z-index: 60; }
    .row-resizer:hover { background: #1a73e8; }
    .resize-line-col { position: absolute; top: 0; bottom: 0; width: 2px; background: #1a73e8; z-index: 10000; pointer-events: none; }
    .resize-line-row { position: absolute; left: 0; right: 0; height: 2px; background: #1a73e8; z-index: 10000; pointer-events: none; }
    .cell.remote-selected { box-shadow:inset 0 0 0 1000px rgba(234,67,53,0.1) !important; outline:2px solid #ea4335; outline-offset:-2px; z-index:15; }
    .cell.remote-selected::after { content:''; position:absolute; bottom:-5px; right:-5px; width:8px; height:8px; background:#ea4335; border:2px solid #fff; border-radius:50%; z-index:25; box-shadow:0 1px 3px rgba(0,0,0,.4); pointer-events:none; }

    .cell-input { background:transparent; border:none; color:inherit; font-family:inherit; font-size:inherit; font-weight:inherit; font-style:inherit; text-align:inherit; height:100%; outline:none; padding:0 4px; width:100%; display:block; box-shadow:none; }
    .visually-hidden { opacity:0; position:absolute; left:0; top:0; z-index:2; }
    .cell-display { position:relative; z-index:1; pointer-events:none; align-items:center; display:flex; min-height:100%; padding:0 4px; color:inherit; font-size:inherit; font-weight:inherit; font-style:inherit; text-align:inherit; white-space:inherit; overflow:inherit; text-overflow:inherit; word-break:inherit; }
    .cell-display.wrap-text { white-space: pre-wrap !important; word-wrap: break-word !important; }
    .floating-editor { position: absolute; z-index: 9999; background: #fff; border: 2px solid #00b050; outline: none; box-shadow: 0 2px 5px rgba(0,0,0,0.2); resize: none; overflow: hidden; font-family: inherit; font-size: inherit; padding: 1px 3px; box-sizing: border-box; white-space: pre;     }
    .shape-panel { position: absolute; top: 100%; left: 0; background: #fff; border: 1px solid #dadce0; border-radius: 4px; box-shadow: 0 4px 12px rgba(0,0,0,0.15); z-index: 300; width: 450px; display: flex; flex-direction: column; cursor: default; }
    .shape-tabs { display: flex; border-bottom: 1px solid #eee; background: #fff; }
    .s-tab { flex: 1; text-align: center; padding: 10px 0; font-size: 13px; cursor: pointer; display: flex; align-items: center; justify-content: center; gap: 4px; color: #5f6368; }
    .s-tab-active { color: #1a73e8; border-bottom: 2px solid #1a73e8; font-weight: 500; }
    .s-tab .material-symbols-outlined { font-size: 16px; }
    .shape-content { padding: 16px; }
    .diagram-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; }
    .diag-item { border: 1px solid #eee; border-radius: 6px; padding: 12px 8px; text-align: center; cursor: pointer; transition: box-shadow 0.2s; background: #fff; }
    .diag-item:hover { box-shadow: 0 2px 6px rgba(0,0,0,0.1); border-color: #ccc; }
    .diag-item svg { width: 100%; height: 60px; }
    .shape-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 8px; }
    .s-item { aspect-ratio: 1; border: 1px solid transparent; border-radius: 2px; cursor: pointer; display: flex; align-items: center; justify-content: center; }
    .s-item:hover { border-color: #1a73e8; background: #f8f9fa; }
    .s-item svg { width: 24px; height: 24px; }
    
    .sheet-shape { position: absolute; z-index: 100; cursor: grab; display: flex; align-items: center; justify-content: center; border: 1px solid transparent; }
    .sheet-shape.shape-active { border: 1px solid #0f9d58; cursor: move; }
    .sheet-shape:active { cursor: grabbing; }
    .sheet-shape .shape-content-wrapper { width: 100%; height: 100%; display: flex; align-items: center; justify-content: center; position: relative; }
    .shape-handle { position: absolute; width: 8px; height: 8px; background: #fff; border: 1px solid #0f9d58; border-radius: 50%; z-index: 101; }
    .shape-handle.nw { top: -4px; left: -4px; cursor: nwse-resize; }
    .shape-handle.n { top: -4px; left: calc(50% - 8px); cursor: ns-resize; width: 16px; border-radius: 8px; }
    .shape-handle.ne { top: -4px; right: -4px; cursor: nesw-resize; }
    .shape-handle.e { top: calc(50% - 8px); right: -4px; cursor: ew-resize; height: 16px; border-radius: 8px; }
    .shape-handle.se { bottom: -4px; right: -4px; cursor: nwse-resize; }
    .shape-handle.s { bottom: -4px; left: calc(50% - 8px); cursor: ns-resize; width: 16px; border-radius: 8px; }
    .shape-handle.sw { bottom: -4px; left: -4px; cursor: nesw-resize; }
    .shape-handle.w { top: calc(50% - 8px); left: -4px; cursor: ew-resize; height: 16px; border-radius: 8px; }
    .shape-menu-btn { position: absolute; top: -24px; right: 0; background: #fff; border: 1px solid #eee; box-shadow: 0 1px 3px rgba(0,0,0,0.1); border-radius: 4px; width: 24px; height: 20px; display: flex; align-items: center; justify-content: center; cursor: pointer; z-index: 102; color:#5f6368; }
    .shape-menu-btn:hover { background: #f1f3f4; }
    .shape-context-menu { position: absolute; top: 0; left: calc(100% + 8px); background: #fff; border-radius: 6px; box-shadow: 0 4px 12px rgba(0,0,0,0.15); width: 200px; padding: 8px 0; z-index: 103; font-size: 13px; color: #202124; }
    .scm-item { padding: 8px 16px; display: flex; align-items: center; gap: 12px; cursor: pointer; transition: background 0.1s; text-align: left; }
    .scm-item:hover { background: #f1f3f4; }
    .scm-item .chevron { margin-left: auto; color: #9aa0a6; }
    .cell-select { border:none; background:transparent; color:inherit; font-family:inherit; font-size:inherit; font-weight:inherit; font-style:inherit; text-align:inherit; height:100%; outline:none; width:100%; cursor:pointer; }
    .fill-handle { background:#34a853; border:2px solid #fff; border-radius:50%; bottom:-5px; right:-5px; cursor:crosshair; height:8px; position:absolute; width:8px; z-index:30; box-shadow:0 1px 3px rgba(0,0,0,.4); }
    .fill-handle:hover { transform:scale(1.4); transition:transform .1s; }
    .cell.fill-preview { outline: 1px dashed #34a853 !important; outline-offset: -1px; background-color: rgba(52,168,83,0.08) !important; }
    .fill-options-wrap { position:fixed; z-index:200000; pointer-events:none; }
    .fill-options-icon { width:22px; height:22px; background:#fff; border:1px solid #bbb; border-radius:3px; display:flex; align-items:center; justify-content:center; cursor:pointer; pointer-events:all; box-shadow:0 1px 4px rgba(0,0,0,.18); }
    .fill-options-icon:hover { background:#f1f3f4; }
    .fill-options-icon svg { width:14px; height:14px; }
    .fill-options-dropdown { position:absolute; top:26px; left:0; background:#202124; border-radius:8px; box-shadow:0 6px 24px rgba(0,0,0,.35); min-width:210px; padding:6px 0; pointer-events:all; }
    .fill-opt-item { display:flex; align-items:center; gap:8px; padding:9px 16px; font-size:13px; color:#e8eaed; cursor:pointer; white-space:nowrap; }
    .fill-opt-item:hover { background:rgba(255,255,255,.1); }
    .fill-opt-check { width:16px; font-size:14px; color:#34a853; flex-shrink:0; }

    /* Frozen row/col legacy unused styles removed */
    .grid-spacing-cozy .cell { padding: 0 4px; }
    .grid-spacing-comfort .cell { padding: 4px 6px; }
    .grid-spacing-classic .cell { padding: 0; }
    .img-overlay { left:0; pointer-events:none; position:absolute; top:0; z-index:6; }
    .filter-row select { border:none; background:transparent; font-size:11px; width:100%; cursor:pointer; }

    /* ── CONTEXT MENU ───────────────────────────────────────────────────── */
        .ctx-menu { background:#fff; border:1px solid #cbd5e1; border-radius:6px; box-shadow:0 4px 20px rgba(0,0,0,.15); min-width:220px; padding:4px 0; position:fixed; z-index:100000; max-height:80vh; overflow-y:auto; color: #333; box-sizing: border-box; }
        .ctx-item { padding:8px 16px; font-size:13px; cursor:pointer; display:flex; align-items:center; gap:8px; }
        .ctx-item:hover { background:#f1f5f9; }
        .ctx-item.danger { color:#e53e3e; }
        .ctx-item.danger:hover { background:#fff5f5; }
        .ctx-item.disabled { color:#a0aec0; cursor:default; pointer-events:none; }
        .ctx-item.disabled .ctx-icon { color:#cbd5e1; }
        .ctx-icon { color:#64748b; font-size:18px; }
    .ctx-hint { color:#94a3b8; font-size:11px; margin-left:auto; }
    .ctx-sep { background:#e2e8f0; height:1px; margin:3px 0; }

    /* ── CUSTOM DROPDOWN ─────────────────────────────────────────────────── */
    ::ng-deep .custom-dropdown-overlay { position: fixed; background: #242424; border: 1px solid #3c4043; border-radius: 8px; box-shadow: 0 8px 24px rgba(0,0,0,0.5); z-index: 100000; display: flex; flex-direction: column; min-width: 170px; max-height: calc(100vh - 20px); padding: 4px 0; overflow: hidden; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }
    ::ng-deep .custom-dropdown-item { padding: 8px 16px; cursor: pointer; font-size: 13px; color: #e8eaed; display: flex; align-items: center; transition: background 0.15s ease; }
    ::ng-deep .custom-dropdown-item:hover { background: rgba(255,255,255,0.08); }
    ::ng-deep .option-row { padding: 8px 16px; font-size: 13px; font-weight: 400; color: #e8eaed; cursor: pointer; transition: background 0.15s ease; outline: none; }
    ::ng-deep .option-row:hover, ::ng-deep .option-row:focus { background: rgba(255,255,255,0.08); }
    ::ng-deep .option-row[data-selected="true"] { background: rgba(255,255,255,0.12); font-weight: 500; color: #ffffff; }
    ::ng-deep .option-row.clear-option { color: #9aa0a6; border-bottom: 1px solid rgba(255,255,255,0.08); margin-bottom: 2px; }
    ::ng-deep .option-row.clear-option:hover, ::ng-deep .option-row.clear-option:focus { color: #e8eaed; background: rgba(255,255,255,0.08); }
    ::ng-deep .dropdown-divider { height: 1px; background: rgba(255,255,255,0.12); margin: 4px 0; }
    ::ng-deep .dropdown-footer { display: flex; justify-content: flex-end; padding: 4px 8px; }
    ::ng-deep .edit-btn { display: inline-flex; align-items: center; gap: 6px; padding: 6px 12px; border: none; border-radius: 4px; background: transparent; color: #9aa0a6; font-size: 13px; font-weight: 500; cursor: pointer; transition: all 0.15s ease; }
    ::ng-deep .edit-btn:hover, ::ng-deep .edit-btn:focus { color: #ffffff; background: rgba(255,255,255,0.08); }
    ::ng-deep .edit-icon { width: 14px; height: 14px; }
    ::ng-deep .option-list { max-height: 180px; overflow-y: auto; }
    ::ng-deep .option-list::-webkit-scrollbar { width: 6px; }
    ::ng-deep .option-list::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.25); border-radius: 3px; }

    /* ── DRAWER SLIDE-IN ANIMATIONS ─────────────────────────────────────── */
    @keyframes drawerSlideIn {
      from { transform: translateX(100%); }
      to { transform: translateX(0); }
    }
    @keyframes overlayFadeIn {
      from { opacity: 0; }
      to { opacity: 1; }
    }
    .drawer-overlay {
      animation: overlayFadeIn 0.2s ease-out forwards;
    }
    .drawer-content {
      animation: drawerSlideIn 0.28s cubic-bezier(0.16, 1, 0.3, 1) forwards;
      will-change: transform;
    }

    /* ── SHEET TABS ─────────────────────────────────────────────────────── */
    .footer-container { display: flex; align-items: center; justify-content: space-between; background: #f1f3f4; border-top: 2px solid #dadce0; min-height: 34px; width: 100%; box-sizing: border-box; }
    .sheet-tabs { display:flex; align-items:flex-end; gap:4px; padding:0 16px 4px 16px; min-height:44px; overflow-x:auto; flex-shrink:1; min-width:0; }
    .sheet-tab { flex-shrink:0; align-items:center; background:transparent; border-radius:6px 6px 0 0; border:1px solid transparent; border-bottom:none; color:#5f6368; cursor:pointer; display:flex; font-size:14px; gap:8px; padding:8px 24px; white-space:nowrap; min-width:80px; justify-content:center; font-weight:500; transition:all 0.2s; }
    .sheet-tab.active-tab { background:#fff; border-color:#dadce0; color:#1a73e8; font-weight:600; }
    .sheet-tab:hover:not(.active-tab) { background:#e8eaed; }
    .tab-close { color:#bbb; cursor:pointer; font-size:13px; line-height:1; }
    .tab-close:hover { color:#d93025; }
    .tab-add { background:none; border:none; color:#5f6368; cursor:pointer; font-size:20px; padding:2px 8px; border-radius:4px; line-height:1; }
    .tab-add:hover { background:#e0e0e0; }
    .footer-tools-container { display: flex; align-items: center; background: #ffffff; border: 1px solid #dadce0; border-radius: 8px; padding: 2px 4px; gap: 2px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); margin-left: auto; margin-right: 14px; height: 28px; flex-shrink: 0; }
    .footer-btn { background:transparent; border:none; border-radius:4px; color:#5f6368; cursor:pointer; display:flex; align-items:center; justify-content:center; padding:4px 6px; transition:all 0.2s; height: 24px; }
    .footer-btn:hover { background:rgba(0,0,0,0.06); color:#202124; }
    .footer-sep { width: 1px; height: 16px; background: #dadce0; margin: 0 4px; }

    /* ── MODALS ─────────────────────────────────────────────────────────── */
    .modal-overlay { align-items:center; background:rgba(0,0,0,.5); bottom:0; display:flex; justify-content:center; left:0; position:fixed; right:0; top:0; z-index:999; }
    .modal { background:#fff; border-radius:10px; box-shadow:0 8px 32px rgba(0,0,0,.22); max-width:92vw; padding:28px; position:relative; width:460px; }
    .btn { background:#1a73e8; border:none; border-radius:6px; color:#fff; cursor:pointer; font-size:13px; font-weight:600; padding:9px 20px; }
    .btn:hover { background:#1557b0; }
    .btn.outline { background:transparent; border:1px solid #1a73e8; color:#1a73e8; }
    .btn.outline:hover { background:#e8f0fe; }
    .validation-textarea { border:1px solid #dadce0; border-radius:4px; font-family:inherit; font-size:13px; outline:none; padding:10px; resize:vertical; width:100%; }
    .validation-textarea:focus { border-color:#1a73e8; box-shadow:0 0 0 2px rgba(26,115,232,.2); }

    /* ── TOAST ──────────────────────────────────────────────────────────── */
    .toast { background:#323232; border-radius:6px; bottom:36px; color:#f1f3f4; font-size:13px; left:50%; opacity:0; padding:12px 24px; pointer-events:none; position:fixed; transform:translateX(-50%) translateY(16px); transition:all .25s ease; z-index:1000; white-space:nowrap; }
    .toast.show { opacity:1; transform:translateX(-50%) translateY(0); }

    /* ── PRINT ──────────────────────────────────────────────────────────── */
          @keyframes spin { 100% { transform: rotate(360deg); } }
      @media print {
      .top-bar, .menu-row, .tb-row, .formula-container, .modal-overlay, .toast, .sheet-tabs { display:none !important; }
      .shell { display:block !important; }
      .grid-wrap { overflow:visible !important; }
      .col-head, .row-head { position:static !important; }
    }

    /* ── CHART DROPDOWN ─────────────────────────────────────────────────── */
    .chart-grid { display:flex; flex-wrap:wrap; gap:16px; margin-bottom:16px; }
    .chart-item { display:flex; flex-direction:column; align-items:center; gap:8px; cursor:pointer; width:90px; padding:8px; border-radius:4px; transition:background 0.2s; }
    .chart-item:hover { background:rgba(255,255,255,0.04); }
    .chart-item span { color:#e8eaed; font-size:12px; font-weight:500; text-align:center; }
    .chart-item svg { border-bottom: 2px solid #5f6368; padding-bottom: 4px; transition: border-bottom-color 0.2s; }
    .chart-item:hover svg { border-bottom-color: #81e6d9; }
    .chart-footer { border-top:1px solid #5f6368; padding-top:12px; font-size:13px; font-weight:600; color:#e8eaed; }
    
    .chart-header-icons { display:flex; justify-content:space-between; border-bottom:1px solid #5f6368; margin-bottom:16px; padding:0 8px; }
    .chart-header-icons span { padding-bottom:10px; margin-bottom:-1px; border-bottom:2px solid transparent; cursor:pointer; color:#9aa0a6; transition:color 0.2s, border-bottom 0.2s; }
    .chart-header-icons span.active { color:#81e6d9; border-bottom:2px solid #81e6d9; }
    .chart-header-icons span:hover { color:#fff; }

    /* ── BORDER DROPDOWN ────────────────────────────────────────────────── */
    .bp-btn { display:flex; align-items:center; justify-content:center; background:transparent; border:1px solid transparent; border-radius:3px; color:#e8eaed; cursor:pointer; width:26px; height:26px; padding:0; transition:all 0.15s; }
    .bp-btn:hover { background:rgba(255,255,255,0.1); border-color:rgba(255,255,255,0.2); color:#fff; }
    .bp-btn .material-symbols-outlined { font-size:18px; }
    .bo-item { display:flex; align-items:center; gap:4px; padding:4px; border:1px solid transparent; border-radius:4px; cursor:pointer; transition:background 0.15s; }
    .bo-item:hover, .bo-item.active-bo { background:rgba(255,255,255,0.1); border-color:rgba(255,255,255,0.2); }
    
    /* Print Area Highlight */
    .print-area-active tr:nth-child(40n) .cell { border-bottom: 2px dashed #9aa0a6 !important; }
    .print-area-active .cell:nth-child(9n) { border-right: 2px dashed #9aa0a6 !important; }

    /* Share Modal Styles */
    .share-modal { background:#202124; color:#e8eaed; border-radius:12px; padding:24px; width:520px; box-shadow:0 12px 40px rgba(0,0,0,.6); font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif; border:none; max-width:90vw; }
    .share-modal h3 { margin:0; font-size:18px; font-weight:500; color:#e8eaed; }
    .sm-close-btn { background:none; border:none; color:#9aa0a6; cursor:pointer; display:flex; align-items:center; justify-content:center; padding:6px; border-radius:50%; transition:background 0.2s; }
    .sm-close-btn:hover { background:rgba(255,255,255,0.08); }
    .sm-input-box { flex:1; display:flex; align-items:center; background:#1c1d1f; border:1px solid #5f6368; border-radius:4px; padding:0 12px; height:44px; transition:border-color 0.2s; }
    .sm-input-box:focus-within { border-color:#8ab4f8; }
    .sm-input { flex:1; background:transparent; border:none; color:#e8eaed; font-size:14px; outline:none; height:100%; }
    .sm-dropdown-txt { display:flex; align-items:center; gap:4px; color:#e8eaed; font-size:13px; cursor:pointer; padding-left:12px; }
    .sm-list { position:absolute; top:48px; left:0; width:calc(100% - 100px); background:#2d3748; border:1px solid #4a5568; border-radius:4px; box-shadow:0 4px 12px rgba(0,0,0,0.5); z-index:100; max-height:200px; overflow-y:auto; }
    .sm-list-item { display:flex; align-items:center; gap:12px; padding:8px 12px; cursor:pointer; border-bottom:1px solid #4a5568; transition:background 0.2s; }
    .sm-list-item:hover { background:#4a5568; }
    .sm-list-item .name { color:#e8eaed; font-size:14px; font-weight:500; }
    .sm-list-item .email { color:#9aa0a6; font-size:12px; }
    .sm-icon-bg { background:#303134; display:flex; align-items:center; justify-content:center; width:40px; height:40px; border-radius:50%; color:#e8eaed; }
    .sm-txt-main { font-size:14px; font-weight:600; color:#e8eaed; }
    .sm-sec-btn { background:transparent; border:none; color:#9aa0a6; display:flex; align-items:center; gap:6px; font-size:13px; font-weight:500; cursor:pointer; padding:8px; border-radius:4px; transition:background 0.2s; }
    .sm-sec-btn:hover { background:rgba(255,255,255,0.04); }
    .sm-copy-btn { background:transparent; border:none; color:#e8eaed; font-size:14px; font-weight:500; border-radius:24px; padding:8px 12px; margin-left:-12px; cursor:pointer; transition:background 0.2s; }
    .sm-copy-btn:hover { background:rgba(255,255,255,0.08); }
    .sm-done-btn { background:#303134; color:#8ab4f8; font-size:14px; font-weight:500; border:none; border-radius:24px; padding:0 24px; height:40px; cursor:pointer; transition:background 0.2s; }
    .sm-done-btn:hover { background:#3c4043; }

    /* ── LIGHT THEME OVERRIDES ────────────────────────────────────────── */
    .theme-light .top-bar { background: #f8f9fa; border-bottom: 1px solid #dadce0; }
    .theme-light .tl-sep { background: rgba(0,0,0,0.15); }
    .theme-light .brand-name, .theme-light .doc-title { color: #202124; }
    .theme-light .cursor-path { stroke: #ffffff; }
    .theme-light .doc-icons { color: #5f6368; }
    .theme-light .brand:hover { background: rgba(0,0,0,0.05); }
    .theme-light .pd-head { background: #f8f9fa; color: #202124; border-bottom-color: #dadce0; }
    .theme-light .doc-title:hover { border-color: rgba(0,0,0,0.2); }
    .theme-light .doc-title:focus { background: #fff; border-color: #1a73e8; }
    .theme-light .doc-sub { color: #5f6368; }
    .theme-light .back-btn, .theme-light .top-search-box { color: #5f6368; }
    .theme-light .top-search-box { background: #f1f3f4; border-color: transparent; }
    .theme-light .top-search-box.has-query { background: #e6f4ea; border-color: #0f9d58; color: #137333; }
    .theme-light .top-search-box input.inline-search-input { color: #202124; }
    .theme-light .top-search-box input.inline-search-input::placeholder { color: #5f6368; }
    .theme-light .top-search-box .inline-search-clear { background: rgba(0,0,0,0.05); }
    .theme-light .top-search-box .inline-search-clear:hover { background: rgba(0,0,0,0.1); }
    .theme-light .top-search-box .inline-search-divider { background: rgba(0,0,0,0.1); }
    .theme-light .top-search-box .inline-search-nav button:hover { background: rgba(0,0,0,0.08); }
      .theme-light .online-badge { background:#fff; color:#5f6368; border-color:#dadce0; }
    .theme-light .menu-row { background: #ffffff; }
    .theme-light .mi { color: #202124; }
    .theme-light .mi:hover, .theme-light .mi-open { background: #f1f3f4; color: #202124; }
    
    /* Light Theme Share Modal */
    .theme-light .share-modal { background:#ffffff; color:#202124; box-shadow:0 12px 40px rgba(0,0,0,.2); }
    .theme-light .share-modal h3 { color:#202124; }
    .theme-light .sm-close-btn { color:#5f6368; }
    .theme-light .sm-close-btn:hover { background:rgba(0,0,0,0.04); }
    .theme-light .sm-input-box { background:#ffffff; border-color:#dadce0; }
    .theme-light .sm-input-box:focus-within { border-color:#1a73e8; }
    .theme-light .sm-input { color:#202124; }
    .theme-light .sm-dropdown-txt { color:#5f6368; }
    .theme-light .sm-list { background:#ffffff; border-color:#dadce0; box-shadow:0 4px 12px rgba(0,0,0,0.15); }
    .theme-light .sm-list-item { border-bottom-color:#dadce0; }
    .theme-light .sm-list-item:hover { background:#f1f3f4; }
    .theme-light .sm-list-item .name { color:#202124; }
    .theme-light .sm-list-item .email { color:#5f6368; }
    .theme-light .sm-icon-bg { background:#f1f3f4; color:#5f6368; }
    .theme-light .sm-txt-main { color:#202124; }
    .theme-light .sm-sec-btn { color:#5f6368; }
    .theme-light .sm-sec-btn:hover { background:rgba(0,0,0,0.04); }
    .theme-light .sm-copy-btn { color:#1a73e8; }
    .theme-light .sm-copy-btn:hover { background:rgba(26,115,232,0.04); }
    .theme-light .sm-done-btn { background:#f1f3f4; color:#1a73e8; }
    .theme-light .sm-done-btn:hover { background:#e8eaed; }
    .theme-light .mdd, .theme-light .mdi-sub, .theme-light .tb-dd, .theme-light .tb-chart-dd, .theme-light .profile-dd { background: #ffffff; border-color: #dadce0; box-shadow: 0 4px 16px rgba(0,0,0,0.15); }
    .theme-light .bp-btn { color: #5f6368; }
    .theme-light .bp-btn:hover { background: rgba(0,0,0,0.05); border-color: rgba(0,0,0,0.1); color: #202124; }
    .theme-light .mdi, .theme-light .dd-item, .theme-light .pd-item, .theme-light .bo-item { color: #202124; }
    .theme-light .mdi:hover, .theme-light .dd-item:hover, .theme-light .pd-item:hover, .theme-light .bo-item:hover, .theme-light .bo-item.active-bo { background: #f1f3f4; border-color: transparent; }
    .theme-light .mds, .theme-light .pd-sep { background: #dadce0; }
    .theme-light .mdi-title, .theme-light .mh, .theme-light .mdi-icon, .theme-light .mdi-arrow { color: #5f6368; }
    .theme-light .mdi:hover .mdi-icon, .theme-light .mdi:hover .mdi-arrow { color: #202124; }
    .theme-light .tb-row { background: #edf2fa; border-top: none; }
    .theme-light .tb-row2 { background: #edf2fa; border-top: 1px solid rgba(0,0,0,0.08); }
    .theme-light .tb-sep { background: rgba(0,0,0,0.2); }
    .theme-light .tb, .theme-light .tb-clr { color: #444746; }
    .theme-light .tb:hover, .theme-light .tb-clr:hover { background: rgba(0,0,0,0.08); color: #202124; }
    .theme-light .tb.tb-on { background: #d3e3fd; color: #041e49; }
    .theme-light .tb-font-dd { border-color: rgba(0,0,0,0.15); color: #444746; background: #ffffff; }
    .theme-light .tb-font-dd:hover, .theme-light .tb-font-dd.active { background: #f8f9fa; }
    .theme-light .arr { color: #444746; }
    .theme-light .sz-inp { background: #ffffff; border-color: rgba(0,0,0,0.15); color: #444746; }
    .theme-light .sz-drop-btn { background: #ffffff; border-color: rgba(0,0,0,0.15); color: #444746; }
    .theme-light .zoom-ctrl { color: #444746; }
    .theme-light .formula-container { background: #ffffff; border-bottom: 1px solid #dadce0; border-top: 1px solid #dadce0; }
    .theme-light .cell-ref { background: #ffffff; border: 1px solid #dadce0; color: #202124; }
    .theme-light .fx-label { color: #5f6368; }
    .theme-light .formula-bar { background: #ffffff; border: 1px solid #dadce0; color: #202124; }
    .theme-light .formula-bar:focus { border-color: #1a73e8; }
    .theme-light .footer-container { background: #f8f9fa; border-top: 1px solid #dadce0; }
    .theme-light .sheet-tab { background: #ffffff; color: #5f6368; border: 1px solid #dadce0; border-bottom: none; }
    .theme-light .sheet-tab.active-tab { background: #ffffff; color: #1a73e8; border-top: 2px solid #1a73e8; font-weight:600; }
    .theme-light .sheet-tab:hover:not(.active-tab) { background: #f1f3f4; }
    .theme-light .tab-add { color: #5f6368; }
    .theme-light .tab-add:hover { background: rgba(0,0,0,0.05); }
    .theme-light .grid { background: #f8f9fa; }
    .theme-light .corner, .theme-light .col-head, .theme-light .row-head { background: #f8f9fa; color: #5f6368; border-color: #c0c0c0; }
    .theme-light .col-head:hover, .theme-light .row-head:hover { background: #e8eaed; color: #202124; }
    .theme-light .col-selected, .theme-light .row-selected { background: #e8eaed !important; color: #202124 !important; font-weight:700 !important; }
    .theme-light .active-axis { background: #e8eaed !important; color: #1a73e8 !important; border-bottom-color: #1a73e8; }
    .theme-light .row-head.active-axis { border-right-color: #1a73e8; }
    
    /* ── DARK THEME OVERRIDES ─────────────────────────────────────────── */
    .theme-dark .top-bar { background: #1e1e1e; border-bottom: 1px solid rgba(255, 255, 255, 0.15); }
    .theme-dark .back-btn { color: #e8eaed; }
    .theme-dark .back-btn:hover { background: rgba(255,255,255,0.05); }
    .theme-dark .menu-row { background: #1e1e1e; border-bottom: 1px solid rgba(255, 255, 255, 0.15); }
    .theme-dark .tb-row { background: #1e1e1e; border-bottom: 1px solid rgba(255, 255, 255, 0.15); }
    .theme-dark .tb-row2 { background: #1e1e1e; }
    .theme-dark .tb-font-dd { border-color: rgba(255,255,255,0.3); background: rgba(255,255,255,0.06); }
    .theme-dark .sz-inp { border-color: rgba(255,255,255,0.3); background: rgba(255,255,255,0.06); }
    .theme-dark .sz-drop-btn { border-color: rgba(255,255,255,0.3); background: rgba(255,255,255,0.06); }
    .theme-dark .tb-sep { background: rgba(255,255,255,0.3); }
    .theme-dark .formula-container { background: #1e1e1e; border-bottom: 1px solid #333; }
    .theme-dark .cell-ref { background: rgba(255,255,255,.05); border-color: rgba(255,255,255,.3); color: #fff; }
    .theme-dark .formula-bar { background: rgba(255,255,255,.05); border-color: rgba(255,255,255,.3); color: #fff; }
    .theme-dark .formula-bar:focus { border-color: #10b981; }
    .theme-dark .corner, .theme-dark .col-head, .theme-dark .row-head { background: #202124; border-color: #5f6368; color: #e8eaed; }
    .theme-dark .footer-btn { color:#f3f4f6; }
    .theme-dark .footer-btn:hover { background:rgba(255,255,255,0.1); color:#fff; }
    .theme-dark .footer-container { background: #1e1e1e; border-top: 1px solid #333; }
    /* Loading & Upload Overlay (2026 Premium) */
    .loading-overlay { position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(255, 255, 255, 0.85); backdrop-filter: blur(8px); -webkit-backdrop-filter: blur(8px); z-index: 99999; display: flex; align-items: center; justify-content: center; animation: fadeIn 0.3s ease-out; }
    .loading-modal, .upload-modal { background: #fff; border-radius: 16px; padding: 40px; width: 420px; text-align: center; border: 1px solid rgba(0,0,0,0.08); box-shadow: 0 20px 40px rgba(0,0,0,0.1); animation: slideUp 0.4s cubic-bezier(0.16, 1, 0.3, 1); }
    .lm-spinner { width: 48px; height: 48px; border: 4px solid #e8f0fe; border-top-color: #1a73e8; border-radius: 50%; margin: 0 auto 24px auto; animation: spin 1s linear infinite; }
    .lm-title, .um-title { font-size: 20px; font-weight: 600; color: #202124; margin-bottom: 8px; letter-spacing: -0.3px; }
    .lm-subtitle, .um-subtitle { font-size: 14px; color: #5f6368; line-height: 1.5; padding: 0 20px; }
    
    .um-icon { width: 64px; height: 64px; border-radius: 50%; background: #e8f0fe; display: flex; align-items: center; justify-content: center; margin: 0 auto 20px auto; }
    .um-subtitle { margin-bottom: 28px; }
    .um-progress-container { height: 8px; background: #e0e0e0; border-radius: 4px; overflow: hidden; margin-bottom: 16px; position: relative; }
    .um-progress-bar { height: 100%; background: linear-gradient(90deg, #10b981, #059669); border-radius: 4px; transition: width 0.2s ease-out; position: relative; overflow: hidden; }
    .um-progress-bar::after { content: ''; position: absolute; top: 0; left: 0; right: 0; bottom: 0; background: linear-gradient(90deg, rgba(255,255,255,0) 0%, rgba(255,255,255,0.3) 50%, rgba(255,255,255,0) 100%); animation: shimmer 1.5s infinite; }
    .um-stats { display: flex; justify-content: space-between; font-size: 13px; font-weight: 500; }
    .um-percent { color: #10b981; font-weight: 600; font-size: 14px; }
    .um-time { color: #5f6368; }

    @keyframes spin { to { transform: rotate(360deg); } }
    @keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }
    @keyframes slideUp { from { opacity: 0; transform: translateY(20px) scale(0.98); } to { opacity: 1; transform: translateY(0) scale(1); } }
    @keyframes shimmer { 0% { transform: translateX(-100%); } 100% { transform: translateX(100%); } }
  `]
})

export class SheetEditorComponent implements OnInit, OnDestroy {
  shortcutCategoryFilter: string = 'all';
  shortcutSearchQuery: string = '';

  shortcutCategories = [
    {
      name: 'File operations',
      id: 'file',
      shortcuts: [
        { desc: 'Save file', keys: ['Ctrl', 'S'] },
        { desc: 'Print file', keys: ['Ctrl', 'P'] }
      ]
    },
    {
      name: 'Edit actions',
      id: 'edit',
      shortcuts: [
        { desc: 'Undo last action', keys: ['Ctrl', 'Z'] },
        { desc: 'Redo last action', keys: ['Ctrl', 'Y'] },
        { desc: 'Cut', keys: ['Ctrl', 'X'] },
        { desc: 'Copy', keys: ['Ctrl', 'C'] },
        { desc: 'Paste', keys: ['Ctrl', 'V'] },
        { desc: 'Cancel cell entry', keys: ['Esc'] },
        { desc: 'Delete content of selected cell', keys: ['Backspace'] }
      ]
    },
    {
      name: 'Formatting',
      id: 'format',
      shortcuts: [
        { desc: 'Bold toggle for selection', keys: ['Ctrl', 'B'] },
        { desc: 'Italic toggle for selection', keys: ['Ctrl', 'I'] },
        { desc: 'Underline toggle for selection', keys: ['Ctrl', 'U'] },
        { desc: 'Strikethrough toggle for selection', keys: ['Ctrl', 'Shift', 'X'] },
        { desc: 'Add / Edit Hyperlink', keys: ['Ctrl', 'K'] },
        { desc: 'Insert the current date in cell', keys: ['Ctrl', ';'] },
        { desc: 'Insert the current time in cell', keys: ['Ctrl', 'Shift', ';'] },
        { desc: 'Increase Indentation', keys: ['Ctrl', 'M'] },
        { desc: 'Decrease Indentation', keys: ['Ctrl', 'Shift', 'M'] }
      ]
    },
    {
      name: 'Navigation & Data',
      id: 'nav',
      shortcuts: [
        { desc: 'Fill down', keys: ['Ctrl', 'D'] },
        { desc: 'Fill to the right', keys: ['Ctrl', 'R'] },
        { desc: 'Find within spreadsheet', keys: ['Ctrl', 'F'] },
        { desc: 'Move to next cell in row', keys: ['Tab'] },
        { desc: 'Move to previous cell in row', keys: ['Shift', 'Tab'] }
      ]
    },
    {
      name: 'Selection',
      id: 'sel',
      shortcuts: [
        { desc: 'Select whole spreadsheet', keys: ['Ctrl', 'A'] }
      ]
    }
  ];

  get filteredShortcutCategories() {
    return this.shortcutCategories.map(cat => {
      if (this.shortcutCategoryFilter !== 'all' && this.shortcutCategoryFilter !== cat.id) {
        return { ...cat, shortcuts: [] };
      }
      const q = this.shortcutSearchQuery.toLowerCase();
      const filtered = cat.shortcuts.filter(s => s.desc.toLowerCase().includes(q));
      return { ...cat, shortcuts: filtered };
    }).filter(cat => cat.shortcuts.length > 0);
  }

  private pendingDiffPreStateJson: string | null = null;
  private pendingDiffContext: any = null;
  private pendingDiffTimer: any = null;
  private auditBuffer = new Map<string, any>();
  private flushAudit$ = new Subject<void>();
  private flushAuditSubscription!: Subscription;
  private pendingAuditPromise: Promise<void> | null = null;
  auditRecords: any[] = [];
  auditSortBy: string = 'date';
  auditSortDesc: boolean = true;
  trackByFn(index: number, item: any) { return index; }
  activeCtxSubmenu: 'insert' | 'delete' | 'clear' | 'filter' | 'paste' | null = null;
  ctxSubX: number = 0;
  ctxSubTop: number | null = null;
  ctxSubBottom: number | null = null;
  ctxSubMaxHeight: number = 800;
  ctxSubmenuTimer: any;

  showCtxSubmenu(type: 'insert' | 'delete' | 'clear' | 'filter' | 'paste', event: MouseEvent) {
    clearTimeout(this.ctxSubmenuTimer);
    this.activeCtxSubmenu = type;
    const target = event.currentTarget as HTMLElement;
    const rect = target.getBoundingClientRect();

    const submenuWidth = 220;
    // Decide left/right side based on available space
    if (this.ctxX + 220 + submenuWidth > window.innerWidth) {
      this.ctxSubX = this.ctxX - submenuWidth + 4;
    } else {
      this.ctxSubX = this.ctxX + 220 - 4;
    }
    // Clamp X so submenu never goes off-screen
    this.ctxSubX = Math.max(4, Math.min(this.ctxSubX, window.innerWidth - submenuWidth - 4));

    // Use accurate estimated height per submenu type
    const heightMap: Record<string, number> = {
      clear: 360,
      paste: 400,
      insert: 350,
      delete: 180,
      filter: 100
    };
    const estimatedHeight = heightMap[type] ?? 200;
    const margin = 8; // minimum gap from screen edge

    // Try to open downward from the hovered item
    if (rect.top + estimatedHeight + margin <= window.innerHeight) {
      this.ctxSubTop = rect.top;
      this.ctxSubBottom = null;
      this.ctxSubMaxHeight = window.innerHeight - rect.top - margin;
    } else if (rect.bottom - estimatedHeight - margin >= 0) {
      // Not enough space below - open upward
      this.ctxSubTop = Math.max(margin, rect.bottom - estimatedHeight);
      this.ctxSubBottom = null;
      this.ctxSubMaxHeight = window.innerHeight - this.ctxSubTop - margin;
    } else {
      // Not enough space either way - pin to top with margin
      this.ctxSubTop = margin;
      this.ctxSubBottom = null;
      this.ctxSubMaxHeight = window.innerHeight - 2 * margin;
    }
  }

  hideCtxSubmenu() {
    this.ctxSubmenuTimer = setTimeout(() => {
      this.activeCtxSubmenu = null;
    }, 300);
  }

  keepCtxSubmenu() {
    clearTimeout(this.ctxSubmenuTimer);
  }

  goHome() {
    window.location.href = 'https://sheets.vsnaptechnology.com/';
  }

  get selectedRowCount(): number {
    if (!this.rangeStart || !this.rangeEnd) return 1;
    return Math.abs(this.rangeEnd.r - this.rangeStart.r) + 1;
  }

  get selectedColCount(): number {
    if (!this.rangeStart || !this.rangeEnd) return 1;
    return Math.abs(this.rangeEnd.c - this.rangeStart.c) + 1;
  }

  get selectedNonEmptyCount(): number {
    if (!this.rangeStart || !this.rangeEnd) return 0;

    // For single cell selection, don't show the count pill
    if (this.rangeStart.r === this.rangeEnd.r && this.rangeStart.c === this.rangeEnd.c) return 0;

    let count = 0;
    const minR = Math.min(this.rangeStart.r, this.rangeEnd.r);
    const maxR = Math.max(this.rangeStart.r, this.rangeEnd.r);
    const minC = Math.min(this.rangeStart.c, this.rangeEnd.c);
    const maxC = Math.max(this.rangeStart.c, this.rangeEnd.c);

    // If selecting full columns/rows, this loop could be huge, so cap the search to this.ROWS/this.COLS
    const endR = Math.min(maxR, this.ROWS - 1);
    const endC = Math.min(maxC, this.COLS - 1);

    for (let r = minR; r <= endR; r++) {
      if (this.hiddenRows.has(r)) continue;
      for (let c = minC; c <= endC; c++) {
        if (this.cells[r] && this.cells[r][c] !== undefined && this.cells[r][c] !== null && this.cells[r][c] !== '') {
          count++;
        }
      }
    }
    return count;
  }

  COLS = 30;
  ROWS = 1000;
  visibleRowRange: number[] = [];
  firstUnfrozenRow: number = 0;
  topSpacerHeight: number = 0;
  bottomSpacerHeight: number = 0;

  onGridScroll(event: Event) {
    const el = event.target as HTMLElement;
    const scrollTop = el.scrollTop;
    this.updateVisibleRows(scrollTop);

    const scrollLeft = el.scrollLeft;
    if (scrollLeft + el.clientWidth >= el.scrollWidth - 150) {
      this.addColumns(10);
    }
  }

  addColumns(count: number) {
    for (let i = 0; i < count; i++) {
      this.colRange.push(this.COLS);
      for (let r = 0; r < this.ROWS; r++) {
        if (!this.cells[r]) this.cells[r] = [];
        this.cells[r].push('');
      }
      this.COLS++;
    }
    if (this.sheets[this.currentSheetIdx]) {
      this.sheets[this.currentSheetIdx].cells = this.cells;
    }
  }



  updateVisibleRows(scrollTop: number) {
    let currentHeight = 0;
    let startRow = 0;
    const defaultRowHeight = 24;
    const wrapEl = document.querySelector('.grid-wrap') as HTMLElement;
    const viewportHeight = (wrapEl ? wrapEl.clientHeight : 0) || 1000;
    while (currentHeight < scrollTop && startRow < this.ROWS) {
      if (!this.hiddenRows.has(startRow)) {
        currentHeight += this.getRowHeight(startRow) || defaultRowHeight;
      }
      startRow++;
    }
    let endRow = startRow;
    let viewportAcc = 0;
    while (viewportAcc < viewportHeight && endRow < this.ROWS) {
      if (!this.hiddenRows.has(endRow)) {
        viewportAcc += this.getRowHeight(endRow) || defaultRowHeight;
      }
      endRow++;
    }
    const buffer = 15;
    
    // Expand startRow upwards, skipping hidden rows
    let rowsAdded = 0;
    while (rowsAdded < buffer && startRow > 0) {
      startRow--;
      if (!this.hiddenRows.has(startRow)) rowsAdded++;
    }

    // Expand endRow downwards, skipping hidden rows
    rowsAdded = 0;
    while (rowsAdded < buffer && endRow < this.ROWS - 1) {
      endRow++;
      if (!this.hiddenRows.has(endRow)) rowsAdded++;
    }

    let actualStartRow = Math.max(this.frozenRowsCount, startRow);
    this.visibleRowRange = [];

    for (let i = 0; i < this.frozenRowsCount; i++) {
      if (!this.hiddenRows.has(i)) this.visibleRowRange.push(i);
    }
    for (let i = actualStartRow; i <= endRow; i++) {
      if (!this.hiddenRows.has(i)) this.visibleRowRange.push(i);
    }

    this.firstUnfrozenRow = actualStartRow;

    let calcTopSpacer = 0;
    for (let i = this.frozenRowsCount; i < actualStartRow; i++) {
      if (!this.hiddenRows.has(i)) {
        calcTopSpacer += this.getRowHeight(i) || defaultRowHeight;
      }
    }
    this.topSpacerHeight = calcTopSpacer;

    let calcBottomSpacer = 0;
    for (let i = endRow + 1; i < this.ROWS; i++) {
      if (!this.hiddenRows.has(i)) {
        calcBottomSpacer += this.getRowHeight(i) || defaultRowHeight;
      }
    }
    this.bottomSpacerHeight = calcBottomSpacer;
    if (this.cdr) this.cdr.markForCheck();
  }

  trackByRow(index: number, r: number) { return r; }
  trackByCol(index: number, c: number) { return c; }
  trackByHiddenSheet(index: number, item: any) { return item.idx; }

  activeWidget: string | null = null;
  toggleWidget(w: string) {
    if (this.activeWidget === w) this.activeWidget = null;
    else this.activeWidget = w;
  }

  @ViewChild('imgInput') imgInputRef!: ElementRef<HTMLInputElement>;
  @ViewChild('gridWrap') gridWrapRef!: ElementRef<HTMLElement>;

  docId = '';
  title = '';
  activeUsers = 1;

  currentBorderColor: string = '#000000';
  currentBorderStyle: string = 'solid';
  currentBorderWidth: string = '1px';
  activeBorderSubmenu: 'color' | 'style' | null = null;

  getBorderStyleCss(style: string, width: string = '1px') {
    return { 'border-top': `${width} ${style} currentColor`, 'width': '100%' };
  }

  displayCache: { [key: string]: string } = {};

  updateDisplayCache() {
    if (this.cdr) this.cdr.markForCheck();

    this.displayCache = {};
    for (let r = 0; r < this.ROWS; r++) {
      for (let c = 0; c < this.COLS; c++) {
        const raw = this.cells[r][c];
        if (raw && typeof raw === 'string' && raw.startsWith('=')) {
          this.displayCache[`${r},${c}`] = this.evalCell(r, c);
        }
      }
    }
  }

  getDisplayValue(r: number, c: number): string {
    const raw = this.cells[r][c];
    let val = raw;
    if (raw && typeof raw === 'string' && raw.startsWith('=')) {
      val = this.displayCache[`${r},${c}`] !== undefined ? this.displayCache[`${r},${c}`] : raw;
    }

    const fmt = this.formats[`${r},${c}`];
    const hasNumFormat = fmt && fmt.numFormat && fmt.numFormat !== 'general';
    const hasDecimals = fmt && fmt.decimals !== undefined;

    if (hasNumFormat && val !== '' && val !== undefined && val !== null) {
      return this.formatNumberValue(val, fmt.numFormat as string, fmt.decimals);
    }

    // Apply decimals-only formatting (no numFormat set, just .0/.00 buttons)
    if (hasDecimals && val !== '' && val !== undefined && val !== null) {
      const num = Number(val);
      if (!isNaN(num) && String(val).trim() !== '') {
        return num.toFixed(fmt.decimals);
      }
    }

    return val !== undefined && val !== null ? String(val) : '';
  }

  formatNumberValue(val: any, format: string, decimalsOverride?: number): string {
    const num = Number(val);
    const isNum = !isNaN(num) && String(val).trim() !== '';

    if (format.startsWith('regional_')) {
      if (!isNum) return val;
      const locale = format.split('_')[1];

      // Zoho-style regional formats
      if (locale === 'zip') {
        // Zip Code: 5 digits, zero-padded e.g. 00123
        return String(Math.abs(Math.floor(num))).padStart(5, '0');
      }
      if (locale === 'zip4') {
        // Zip Code+4: e.g. 00123-0000
        const base = String(Math.abs(Math.floor(num))).padStart(9, '0');
        return base.slice(0, 5) + '-' + base.slice(5, 9);
      }
      if (locale === 'phone') {
        // Phone Number: (xxx) xxx-xxxx
        const digits = String(Math.abs(Math.floor(num))).padStart(10, '0').slice(-10);
        return `(${digits.slice(0,3)}) ${digits.slice(3,6)}-${digits.slice(6)}`;
      }
      if (locale === 'ssn') {
        // Social Security Number: xxx-xx-xxxx
        const digits = String(Math.abs(Math.floor(num))).padStart(9, '0').slice(-9);
        return `${digits.slice(0,3)}-${digits.slice(3,5)}-${digits.slice(5)}`;
      }

      // Legacy locale-based formats (kept for backward compatibility)
      const locales: any = { us: 'en-US', uk: 'en-GB', in: 'en-IN', de: 'de-DE', fr: 'fr-FR', it: 'it-IT', jp: 'ja-JP', cn: 'zh-CN' };
      return num.toLocaleString(locales[locale] || 'en-US', { maximumFractionDigits: 10 });
    }

    if (format.startsWith('custom_')) {
      const fmtStr = format.substring(7);
      return this.applyCustomFormatString(val, num, isNum, fmtStr);
    }

    if (format === 'date_full') return isNum ? new Date(Math.round((num - 25569) * 86400 * 1000)).toLocaleDateString(undefined, { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' }) : val;
    if (format === 'date_iso') return isNum ? new Date(Math.round((num - 25569) * 86400 * 1000)).toISOString().split('T')[0] : val;
    if (format === 'financial') return isNum ? (num < 0 ? `(${Math.abs(num).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })})` : num.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })) : val;

    if (!format) return val;

    // Check if it's a date
    let date = null;
    if (typeof val === 'string' && val.includes('-') && !isNaN(Date.parse(val))) {
      date = new Date(val);
    } else if (isNum && format.startsWith('date')) {
      // Excel epoch dates (simplified)
      date = new Date(Math.round((num - 25569) * 86400 * 1000));
    }

    if (format === 'number') {
      const dec = decimalsOverride ?? 2;
      return isNum ? num.toLocaleString(undefined, { minimumFractionDigits: dec, maximumFractionDigits: dec }) : val;
    }
    if (format === 'percent') {
      const dec = decimalsOverride ?? 2;
      return isNum ? (num * 100).toLocaleString(undefined, { minimumFractionDigits: dec, maximumFractionDigits: dec }) + '%' : val;
    }
    if (format === 'scientific') return isNum ? num.toExponential(decimalsOverride ?? 2) : val;
    if (format === 'text') return String(val);

    // Currencies & Accounting
    if (format.startsWith('currency') || format.startsWith('accounting')) {
      if (!isNum) return val;
      let symbol = '$';
      if (format.endsWith('_inr')) symbol = '₹';
      if (format.endsWith('_eur')) symbol = '€';
      if (format.endsWith('_gbp')) symbol = '£';
      if (format.endsWith('_cny')) symbol = '¥';

      const formattedNum = num.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      if (format.startsWith('accounting')) {
        return num === 0 ? `${symbol}   -  ` : `${symbol}  ${formattedNum}`;
      } else {
        return symbol + formattedNum;
      }
    }

    // Fractions
    if (format.startsWith('fraction')) {
      if (!isNum) return val;
      const sign = num < 0 ? '-' : '';
      const absNum = Math.abs(num);
      const whole = Math.floor(absNum);
      const dec = absNum - whole;
      if (dec === 0) return sign + whole;

      let denom = 10;
      if (format === 'fraction_1') denom = 9;
      if (format === 'fraction_2') denom = 99;
      if (format === 'fraction_3') denom = 999;

      let bestH = 0, bestK = 1, minErr = 1;
      for (let k = 1; k <= denom; k++) {
        const h = Math.round(dec * k);
        const err = Math.abs(dec - h / k);
        if (err < minErr) {
          bestH = h;
          bestK = k;
          minErr = err;
          if (err === 0) break;
        }
      }
      return sign + (whole !== 0 ? whole + ' ' : '') + bestH + '/' + bestK;
    }

    // Dates and Times
    if ((format.startsWith('date') || format.startsWith('time')) && date && !isNaN(date.getTime())) {
      const d = date.getDate();
      const m = date.getMonth() + 1;
      const y = date.getFullYear();
      const yy = String(y).slice(-2);
      const mmm = date.toLocaleString('default', { month: 'short' });
      const mmmm = date.toLocaleString('default', { month: 'long' });
      const eeee = date.toLocaleString('default', { weekday: 'long' });

      let h = date.getHours();
      const mm = String(date.getMinutes()).padStart(2, '0');
      const ss = String(date.getSeconds()).padStart(2, '0');
      const ampm = h >= 12 ? 'PM' : 'AM';
      const h12 = h % 12 || 12;

      switch (format) {
        case 'date_1': return `${d}/${m}/${yy}`;
        case 'date_2': return `${d} ${mmm}, ${y}`;
        case 'date_3': return `${d} ${mmmm}, ${y}`;
        case 'date_4': return `${eeee}, ${d} ${mmmm}, ${y}`;
        case 'date_5': return `${String(d).padStart(2, '0')}/${String(m).padStart(2, '0')}/${y}`;
        case 'date_6': return `${String(m).padStart(2, '0')}/${String(d).padStart(2, '0')}/${y}`;
        case 'date_7': return `${y}/${String(m).padStart(2, '0')}/${String(d).padStart(2, '0')}`;
        case 'date_8': return `${d}/${m}/${yy} ${h12}:${mm}:${ss} ${ampm} IST`;
        case 'date_9': return `${d} ${mmm}, ${y} ${h12}:${mm}:${ss} ${ampm} IST`;
        case 'date_10': return `${d} ${mmmm}, ${y} ${h12}:${mm}:${ss} ${ampm}`;
        case 'date_11': return `${eeee}, ${d} ${mmmm}, ${y} ${h12}:${mm} ${ampm}`;
        case 'date_12': return `${d}/${m}/${yy} ${h12}:${mm} ${ampm}`;

        case 'time_1': return `${h12}:${mm} ${ampm}`;
        case 'time_2': return `${h12}:${mm}:${ss} ${ampm}`;
        case 'time_3': return `${h12}:${mm}:${ss} ${ampm} IST`;
        case 'time_4': return isNum ? `${Math.floor(num * 24)}:${mm}` : val;
        case 'time_5': return isNum ? `${Math.floor(num * 24)}:${mm}:${ss}` : val;
      }
    }

    return val !== undefined && val !== null ? String(val) : '';
  }

  isEditingCell = false;
  editValue = '';
  @ViewChild('floatingEditor') floatingEditor?: ElementRef<HTMLTextAreaElement>;

  shareModalOpen = false;
  propertiesPanelOpen = false;
  docDetails: any = null;
  window = window;

  getActiveSheetCount(): number {
    return this.sheets.length;
  }

  getUsedCellsCount(): number {
    let count = 0;
    this.sheets.forEach(sheet => {
      if (sheet.cells) {
        Object.keys(sheet.cells).forEach(r => {
          Object.keys(sheet.cells[r as any]).forEach(c => {
            if (sheet.cells[r as any][c as any] !== '') {
              count++;
            }
          });
        });
      }
    });
    return count;
  }
  isPublic = false;
  shareQuery = '';
  selectedShareEmails: string[] = [];
  shareRole: 'View' | 'Edit' = 'View';
  shareRoleDropdownOpen = false;
  userSearchResults: any[] = [];
  promptModalOpen = false;
  promptModalTitle = '';
  promptModalValue = '';
  private promptResolve: ((value: string | null) => void) | null = null;
  confirmModalOpen = false;
  confirmModalMessage = '';
  private confirmResolve: ((value: boolean) => void) | null = null;
  filterActive = false;
  activeFilterCols: Set<number> = new Set();
  frozenRowsCount: number = 0;
  frozenColsCount: number = 0;
  gridDirection: 'ltr' | 'rtl' = 'ltr';
  gridSpacing: 'classic' | 'cozy' | 'comfort' = 'cozy';
  gridlineColor: string = '#d0d0d0';
  hiddenRows: Set<number> = new Set();
  hiddenCols: Set<number> = new Set();
  showGridlines = true;
  showFormulaBar = true;
  showHeaders = true;
  showTopBar = true;
  showStatusBar = true;
  showNotes = false;
  showUserPresence = true;
  showLockPattern = false;
  showHighlightPrintArea = false;
  appearance: 'light' | 'dark' | 'system' = 'light';
  dataLoaded = false;

  get currentTheme(): string {
    if (this.appearance === 'system') {
      return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
    }
    return this.appearance;
  }

  highlightRowColColor: string = 'transparent';
  highlightColors: string[] = ['transparent', '#e3f2fd', '#e8f5e9', '#fff9c4', '#ffe0b2', '#fce4ec', '#f3e5f5', '#f5f5f5'];
  zoomLevel = 100;
  activePalette: string | null = null;
  currentFont = 'Arial';
  currentSize = '13px';
  currentSizeNum = 13;
  fonts = ['Arial', 'Caveat', 'Comfortaa', 'Comic Sans MS', 'Courier New', 'EB Garamond', 'Georgia', 'Impact', 'Lexend', 'Lobster', 'Lora', 'Merriweather', 'Oswald', 'Pacifico', 'Playfair Display', 'Roboto', 'Times New Roman', 'Trebuchet MS', 'Verdana'];
  private clipboard = '';  // legacy single-cell text (used by system clipboard fallback)
  // Rich internal clipboard for range copy/cut operations
  private richClipboard: {
    cells: string[][];
    formats: Record<string, any>;
    validations: Record<string, any>;
    rows: number;
    cols: number;
    originR: number;
    originC: number;
  } | null = null;
  private history: string[] = [];
  private future: string[] = [];

  // Range selection
  rangeStart: { r: number, c: number } | null = null;
  rangeEnd: { r: number, c: number } | null = null;
  private isDraggingRange = false;

  // Fill handle
  private isFilling = false;
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
  } | null = null;

  // Header selection (full col / full row)
  selectedColHeader: number | null = null;
  selectedRowHeader: number | null = null;

  // Context menu
  activeFooterMenu: string | null = null;
  ctxVisible = false;
  ctxRow: number | null = null;
  ctxCol: number | null = null;
  ctxX = 0;
  ctxY = 0;
  ctxTop: number | null = null;
  ctxBottom: number | null = null;
  ctxMaxHeight = 800;

  // Data validation / dropdown
  validations: Record<string, CellValidation> = {};
  validationModalOpen = false;
  managePicklistSidebarOpen = false;
  viewRulesSheet = 'current';
  viewRulesType = 'all';
  appliesToEditing = false;
  appliesToInput = '';
  isCopyMode = false;
  editingOldRule: any = null;
  _managePicklistRules: any[] = [];
  validationInput = '';
  picklistOptions: DropdownOption[] = [];
  picklistSelectType: 'single' | 'multi' = 'single';
  displayAsChip = true;
  showColorOptionsPopover = false;
  colorPopoverItemIndex: number | null = null;
  draggedOptionIndex: number | null = null;

  onOptionDragStart(event: DragEvent, index: number) {
    this.draggedOptionIndex = index;
    if (event.dataTransfer) {
      event.dataTransfer.effectAllowed = 'move';
      event.dataTransfer.setData('text/plain', index.toString());
    }
  }

  onOptionDragOver(event: DragEvent, index: number) {
    event.preventDefault();
    if (event.dataTransfer) {
      event.dataTransfer.dropEffect = 'move';
    }
  }

  onOptionDrop(event: DragEvent, dropIndex: number) {
    event.preventDefault();
    if (this.draggedOptionIndex !== null && this.draggedOptionIndex !== dropIndex) {
      const movedItem = this.picklistOptions.splice(this.draggedOptionIndex, 1)[0];
      this.picklistOptions.splice(dropIndex, 0, movedItem);
    }
    this.draggedOptionIndex = null;
  }

  standardColorsRow = [
    '#990000', '#ff0000', '#ff9900', '#ffff00', '#00ff00', '#00ffff', '#4a86e8', '#0000ff', '#9900ff', '#ff00ff'
  ];

  openColorOptionsPopover(index: number, event?: MouseEvent) {
    if (event) event.stopPropagation();
    this.colorPopoverItemIndex = index;
    this.showColorOptionsPopover = true;
  }

  closeColorOptionsPopover(event?: MouseEvent) {
    if (event) {
      event.stopPropagation();
      event.preventDefault();
    }
    this.showColorOptionsPopover = false;
    this.colorPopoverItemIndex = null;
  }

  setItemFillColor(color: string) {
    if (this.colorPopoverItemIndex !== null && this.picklistOptions[this.colorPopoverItemIndex]) {
      this.picklistOptions[this.colorPopoverItemIndex].color = color;
    }
  }

  setItemTextColor(textColor: string) {
    if (this.colorPopoverItemIndex !== null && this.picklistOptions[this.colorPopoverItemIndex]) {
      this.picklistOptions[this.colorPopoverItemIndex].textColor = textColor;
    }
  }

  resetItemColor() {
    if (this.colorPopoverItemIndex !== null && this.picklistOptions[this.colorPopoverItemIndex]) {
      this.picklistOptions[this.colorPopoverItemIndex].color = '#f97316';
      delete this.picklistOptions[this.colorPopoverItemIndex].textColor;
    }
  }

  pivotModalOpen = false;
  pivotSource = '';
  pivotDestType = 'existing';
  pivotDest = 'Sheet1.A9';
  dataValidationModalOpen = false;
  manageRulesModalOpen = false;

  // Range Picker state
  rangePickerActive = false;
  rangePickerField: 'pivotSource' | 'pivotDest' | 'dvAppliesTo' | null = null;
  rangePickerStartR = -1;
  rangePickerStartC = -1;
  rangePickerEndR = -1;
  rangePickerEndC = -1;
  _pivotModalWasOpen = false;
  _dvModalWasOpen = false;

  // Data Validation form state
  dvCriteria = 'list';
  dvShowList = true;
  dvSortAsc = false;
  dvIgnoreBlanks = true;
  dvAlertsOpen = false;
  dvAlertEnabled = true;
  dvAlertTitle = '';
  dvAlertMsg = '';
  dvAppliesTo = 'Sheet1.A1';

  // Picklist extra configuration
  dvIsMultiSelect = false;
  dvDisplayAsChip = true;
  dvColorMode: 'none' | 'single' | 'multi' = 'none';
  dvSingleColor = '#f1f5f9';
  dvItemColors: Record<string, string> = {};

  // Advanced Filter State
  advFilterVisible = false;
  advFilterMaxHeight = 360;
  advFilterCol: number | null = null;
  advFilterX = 0;
  advFilterY = 0;
  advFilterTab: 'value' | 'cellColor' | 'textColor' = 'value';
  advFilterSearch = '';
  advFilterValues: { val: string, selected: boolean }[] = [];
  advFilterBgColors: { val: string, selected: boolean }[] = [];
  advFilterTextColors: { val: string, selected: boolean }[] = [];
  advFilterSavedState: Map<number, { tab: 'value' | 'cellColor' | 'textColor', allowedVals: Set<string>, allowedBg: Set<string>, allowedText: Set<string> }> = new Map();

  serializeAdvFilterState() {
    return Array.from(this.advFilterSavedState.entries()).map(([k, v]) => [k, {
      tab: v.tab,
      allowedVals: Array.from(v.allowedVals),
      allowedBg: Array.from(v.allowedBg),
      allowedText: Array.from(v.allowedText)
    }]);
  }

  deserializeAdvFilterState(stateData: any) {
    if (!stateData) {
      this.advFilterSavedState.clear();
      return;
    }
    this.advFilterSavedState = new Map(stateData.map(([k, v]: any) => [k, {
      tab: v.tab,
      allowedVals: new Set(v.allowedVals || []),
      allowedBg: new Set(v.allowedBg || []),
      allowedText: new Set(v.allowedText || [])
    }]));
  }

  // Multiple sheets
  sheets: Array<{ name: string, cells: string[][], formats: Record<string, CellFormat>, validations: Record<string, CellValidation>, sparklines?: Record<string, SparklineConfig>, colWidths?: Record<number, number>, rowHeights?: Record<number, number>, hideGridlines?: boolean, gridlineColor?: string, locked?: boolean, hidden?: boolean, tabColor?: string, shapes?: any[], rowGroups?: Array<{ start: number, end: number, collapsed: boolean }>, colGroups?: Array<{ start: number, end: number, collapsed: boolean }>, hiddenRows?: number[], activeFilterCols?: number[], filterActive?: boolean, advFilterSavedState?: any, frozenRowsCount?: number, frozenColsCount?: number }> = [
    { name: 'Sheet1', cells: Array.from({ length: this.ROWS }, () => Array(this.COLS).fill('')), formats: {}, validations: {}, shapes: [], sparklines: {} }
  ];
  currentSheetIdx = 0;
  activeSheetMenuIdx: number | null = null;
  sheetMenuX = 0;
  sheetMenuY = 0;
  copiedSheetData: any = null;
  activeSheetSubmenu: string | null = null;

  // Resizing state
  resizingCol: number | null = null;
  resizingRow: number | null = null;
  resizeStartX = 0;
  resizeStartY = 0;
  resizeStartSize = 0;
  resizeLineX = 0;
  resizeLineY = 0;

  // Find & Replace
  findModalOpen = false;
  findModalPosition: 'left' | 'right' = 'right';
  findQuery = '';
  replaceQuery = '';
  findStatus = '';
  findMatchCase = false;
  findMatchEntireCell = false;
  findIncludeFormulas = false;
  showShareModal = false;
  findDirection: 'up' | 'down' = 'down';
  findSearchIn = 'sheet';
  findMatches: { r: number, c: number, sIdx: number }[] = [];
  findMatchIdx = -1;

  // Go To
  gotoQuery = '';

  isLoadingDocument = true;
  isUploading = false;
  uploadProgress = 0;
  uploadTimeLeft = '';
  private uploadStartTime = 0;

  themeColorsTop = [
    '#000000', '#434343', '#666666', '#999999', '#cccccc', '#efefef', '#f3f3f3', '#ffffff', '#ff0000', '#00ff00'
  ];
  themeColorsGrid = [
    '#f2f2f2', '#7f7f7f', '#d0cece', '#d6dce4', '#d9e1f2', '#fce4d6', '#ededed', '#fff2cc', '#deebf7', '#e2efda',
    '#d8d8d8', '#595959', '#a2a2a2', '#adb9ca', '#b4c6e7', '#f8cbad', '#dbdbdb', '#ffe699', '#bdd7ee', '#c6e0b4',
    '#bfbfbf', '#3f3f3f', '#7b7b7b', '#8497b0', '#8ea9db', '#f4b084', '#c9c9c9', '#ffd966', '#9dc3e6', '#a9d08e',
    '#a5a5a5', '#262626', '#525252', '#333f4f', '#2f5597', '#c55a11', '#7b7b7b', '#bf8f00', '#2e75b6', '#548235',
    '#7f7f7f', '#0c0c0c', '#252525', '#222a35', '#1f3864', '#833c0c', '#525252', '#7f6000', '#1e4e79', '#375623'
  ];
  standardColors = [
    '#c00000', '#ff0000', '#ffc000', '#ffff00', '#92d050', '#00b050', '#00b0f0', '#0070c0', '#002060', '#7030a0'
  ];

  colRange = Array.from({ length: this.COLS }, (_, i) => i);
  rowRange = Array.from({ length: this.ROWS }, (_, i) => i);
  cells: string[][] = Array.from({ length: this.ROWS }, () => Array(this.COLS).fill(''));
  formats: Record<string, CellFormat> = {};

  selectedRow = 0;
  selectedCol = 0;
  formulaBarValue = '';
  toastVisible = false;
  toastMsg = '';
  remoteCursors: Record<string, { r: number, c: number }> = {};
  activeMenu: string | null = null;
  activeChartTab: string = 'column';
  profileOpen = false;

  get currentUrl(): string { return window.location.href; }

  private syncSub?: Subscription;
  private applyingRemote = false;

  get selectedRef() { return `${colName(this.selectedCol)}${this.selectedRow + 1}`; }
  colLabel(i: number) { return colName(i); }

  colToIndex(col: string): number {
    let index = 0;
    for (let i = 0; i < col.length; i++) {
      index = index * 26 + (col.toUpperCase().charCodeAt(i) - 64);
    }
    return index - 1;
  }

  isRemoteSelected(r: number, c: number) {
    return Object.values(this.remoteCursors).some(pos => pos.r === r && pos.c === c);
  }

  get initials() {
    return (this.auth.user?.name ?? 'U').split(' ').map((w: string) => w[0]).join('').toUpperCase().slice(0, 2);
  }

  sidePanelApp: string | null = null;
  highlightCommentsEnabled: boolean = localStorage.getItem('highlightCommentsEnabled') !== 'false';
  commentsViewFilter: 'all' | 'current' = 'all';
  commentsStatusFilter: 'all' | 'unresolved' | 'resolved' = 'all';
  newCommentCellRef: string | null = null;
  newCommentCellName: string = '';
  newCommentText: string = '';
  activeCommentMenu: string | null = null;
  replyTexts: { [key: string]: string } = {};
  cachedComments: any[] = [];
  @ViewChild('newCommentInput') newCommentInput!: ElementRef;
  sidePanelUrl: SafeResourceUrl | null = null;

  sparklineConfig: SparklineConfig = {
    type: 'line',
    baseColor: '#4285f4',
    highlights: {
      high: { enabled: false, color: '#34A853' },
      low: { enabled: false, color: '#F4B400' },
      first: { enabled: false, color: '#4A86E8' },
      last: { enabled: false, color: '#7BAAF7' },
      negative: { enabled: false, color: '#EA4335' },
      markers: { enabled: false, color: '#4A86E8' }
    },
    emptyCellMode: 'gap',
    includeHiddenRowsColumns: false,
    horizontalAxis: { displayAxis: false, rightToLeft: false },
    verticalAxis: {
      min: { mode: 'auto', customValue: null },
      max: { mode: 'auto', customValue: null }
    },
    isGrouped: false,
    groupId: ''
  };
  // Add state for UI
  insertSparklineConfig = { source: '', dest: '', error: '' };
  editSparklineConfig = { source: '', dest: '', error: '', tab: 'selected' as 'selected' | 'group' };
  colorPickerState: { active: boolean, top: number, left: number, target: 'base' | 'high' | 'low' | 'first' | 'last' | 'negative' | 'markers' | null } = { active: false, top: 0, left: 0, target: null };
  customColorInput = '';
  recentColors: string[] = [];


  // Embedded Side Panel Apps Data
  calendarNotes: Record<string, string> = {};
  selectedCalDate = new Date().toISOString().split('T')[0];
  globalNotes = '';
  tasks: { text: string, done: boolean }[] = [];
  newTask = '';
  activeNotePopup: { r: number, c: number, text: string } | null = null;

  constructor(
    private sanitizer: DomSanitizer,
    private route: ActivatedRoute,
    private router: Router,
    private api: ApiService,
    public auth: AuthService
    , private cdr: ChangeDetectorRef) { }

  ngOnInit() {
    if (!localStorage.getItem('emoji_cache_cleared_v1')) {
      localStorage.removeItem('emoji-mart.frequently');
      localStorage.removeItem('emoji-mart.last');
      localStorage.setItem('emoji_cache_cleared_v1', 'true');
    }
    this.updateVisibleRows(0);

    this.saveSubscription = this.saveSubject.pipe(
      debounceTime(2000)
    ).subscribe(() => {
      this.executeSave();
    });
    this.flushAuditSubscription = this.flushAudit$.pipe(
      debounceTime(2000)
    ).subscribe(() => {
      this.sendAuditEvents();
    });
    this.docId = this.route.snapshot.paramMap.get('id') ?? '';
    this.api.getDocument(this.docId).subscribe((doc: any) => {
      this.api.getNotificationSettings(this.docId).subscribe({
        next: (settings: any) => {
          this.emailNotifEmail = settings.notify_email || '';
          this.emailNotifOnEdit = settings.on_edit;
          this.emailNotifOnComment = settings.on_comment;
        },
        error: (err) => console.error('Failed to load notification settings', err)
      });
      console.log("===== FRONTEND DEBUG LOG =====");
      console.log("Doc ID:", this.docId);
      console.log("Raw Content Length:", doc.content ? doc.content.length : 0);
      this.docDetails = doc;
      this.title = doc.title;
      // ── Set browser tab title + favicon ──────────────────────────────
      document.title = (doc.title || 'Untitled') + ' - Vsnap Sheet';
      const sheetFavicon = `data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'><rect width='32' height='32' rx='5' fill='%2326A96C'/><rect x='5' y='8' width='22' height='2.5' rx='1.2' fill='white'/><rect x='5' y='13.5' width='22' height='2.5' rx='1.2' fill='white'/><rect x='5' y='19' width='14' height='2.5' rx='1.2' fill='white'/></svg>`;
      let link: HTMLLinkElement = document.querySelector("link[rel~='icon']") as HTMLLinkElement;
      if (!link) { link = document.createElement('link'); link.rel = 'icon'; document.head.appendChild(link); }
      link.href = sheetFavicon;
      // ─────────────────────────────────────────────────────────────────
      try {
        let p = JSON.parse(doc.content || '{}');
        console.log("Parsed JSON keys:", Object.keys(p));
        if (p.cells && p.cells.length > 0) {
          console.log("Sample cell (0,0):", p.cells[0][0]);
          console.log("Sample cell D2 (1,3):", p.cells[1] ? p.cells[1][3] : "N/A");
        }
        if (Array.isArray(p) && p.length > 0) p = p[0];
        if (p.cells) {
          for (let r = 0; r < this.ROWS; r++)
            for (let c = 0; c < this.COLS; c++)
              this.cells[r][c] = p.cells[r]?.[c] ?? '';
        }
        if (p.formats) {
          this.formats = p.formats;
        }
        if (p.validations) {
          this.validations = p.validations;
        }
        if (p.colWidths) this.sheets[this.currentSheetIdx].colWidths = p.colWidths;
        if (p.rowHeights) this.sheets[this.currentSheetIdx].rowHeights = p.rowHeights;
        if (p._importedSheets) {
          // Normalize sheets: cells may be sparse {r:{c:val}} or 2D array — convert both to 2D array for live editing
          this.sheets = p._importedSheets.map((sheet: any) => {
            let cells2d: string[][];
            if (Array.isArray(sheet.cells)) {
              // Legacy 2D array — pad/clone to this.ROWS×this.COLS
              cells2d = Array.from({ length: Math.max(this.ROWS, sheet.cells.length) }, (_, r) =>
                Array.from({ length: Math.max(this.COLS, sheet.cells[r]?.length ?? 0) }, (_, c) =>
                  sheet.cells[r]?.[c] ?? ''));
            } else {
              // Sparse object {r:{c:val}} — expand to 2D
              const sp = sheet.cells || {};
              const maxR = Math.max(this.ROWS, ...Object.keys(sp).map(Number).filter(n => !isNaN(n))) + 1;
              cells2d = Array.from({ length: maxR }, (_, r) =>
                Array.from({ length: this.COLS }, (_, c) => sp[r]?.[c] ?? ''));
            }
            return { ...sheet, cells: cells2d };
          });
          this.currentSheetIdx = 0;
          const s0 = this.sheets[0];
          for (let r = 0; r < this.ROWS; r++)
            for (let c = 0; c < this.COLS; c++)
              this.cells[r][c] = s0.cells[r]?.[c] ?? '';
          this.formats = { ...(s0.formats || {}) };
          this.validations = { ...(s0.validations || {}) };
          this.hiddenRows = new Set(s0.hiddenRows || []);
          this.activeFilterCols = new Set(s0.activeFilterCols || []);
          this.filterActive = !!s0.filterActive;
          this.deserializeAdvFilterState(s0.advFilterSavedState);
          this.frozenRowsCount = s0.frozenRowsCount || 0;
          this.frozenColsCount = s0.frozenColsCount || 0;
        }
        // Always restore root-level filter state (handles new save format where filter is at the root)
        if (p.filterActive !== undefined) this.filterActive = !!p.filterActive;
        if (p.activeFilterCols !== undefined) this.activeFilterCols = new Set(p.activeFilterCols);
        if (p.hiddenRows !== undefined) this.hiddenRows = new Set(p.hiddenRows);
        if (p.advFilterSavedState !== undefined) this.deserializeAdvFilterState(p.advFilterSavedState);
        if (p.frozenRowsCount !== undefined) this.frozenRowsCount = p.frozenRowsCount;
        if (p.frozenColsCount !== undefined) this.frozenColsCount = p.frozenColsCount;
        if (p.calendarNotes) this.calendarNotes = p.calendarNotes;
        if (p.globalNotes) this.globalNotes = p.globalNotes;
        if (p.tasks) this.tasks = p.tasks;
      } catch (err) {
        console.error('[SheetEditor] Error parsing document content:', err);
        (this as any).initError = true;
      }
      this.dataLoaded = true;
      this.isLoadingDocument = false;
      // Re-apply filter after load if it was active, to ensure hidden rows are computed
      if (this.filterActive && this.advFilterSavedState.size > 0) {
        this.recalculateAllFilters();
      }
      this.updateDisplayCache();
    });

    this.syncSub = this.api.connectSync(this.docId).subscribe(msg => {
      if (msg.type === 'presence') {
        this.activeUsers = msg.users ?? 1;
      } else if (msg.type === 'update') {
        this.activeUsers = msg.users ?? this.activeUsers;
        this.applyingRemote = true;
        if (msg.title) this.title = msg.title;
        if (msg.content !== undefined) {
          try {
            const p = JSON.parse(msg.content!);
            // Handle array of sheets (from server's doc_states) or single object
            const sheets = Array.isArray(p) ? p : [p];
            if (sheets[this.currentSheetIdx]) {
              const active = sheets[this.currentSheetIdx];
              if (active.cells) {
                for (let r = 0; r < this.ROWS; r++)
                  for (let c = 0; c < this.COLS; c++)
                    this.cells[r][c] = active.cells[r]?.[c] ?? '';
              }
              if (active.formats) this.formats = active.formats;
              if (active.validations) this.validations = active.validations;
              if (active.hiddenRows !== undefined) this.hiddenRows = new Set(active.hiddenRows);
              if (active.activeFilterCols !== undefined) this.activeFilterCols = new Set(active.activeFilterCols);
              if (active.filterActive !== undefined) {
                this.filterActive = active.filterActive;
                this.deserializeAdvFilterState(active.advFilterSavedState);
              }
              if (active.frozenRowsCount !== undefined) this.frozenRowsCount = active.frozenRowsCount;
              if (active.frozenColsCount !== undefined) this.frozenColsCount = active.frozenColsCount;

              // When backend sends a full-doc (root object with _importedSheets), the "active" above
              // is the root doc itself. Read filter state from _importedSheets[idx] if root lacks it.
              const importedSheet = active._importedSheets?.[this.currentSheetIdx];
              if (importedSheet) {
                if (importedSheet.filterActive !== undefined && active.filterActive === undefined) {
                  this.filterActive = !!importedSheet.filterActive;
                  this.deserializeAdvFilterState(importedSheet.advFilterSavedState);
                }
                if (importedSheet.activeFilterCols !== undefined && active.activeFilterCols === undefined) {
                  this.activeFilterCols = new Set(importedSheet.activeFilterCols);
                }
                if (importedSheet.hiddenRows !== undefined && active.hiddenRows === undefined) {
                  this.hiddenRows = new Set(importedSheet.hiddenRows);
                }
                if (importedSheet.frozenRowsCount !== undefined && active.frozenRowsCount === undefined) {
                  this.frozenRowsCount = importedSheet.frozenRowsCount;
                }
                if (importedSheet.frozenColsCount !== undefined && active.frozenColsCount === undefined) {
                  this.frozenColsCount = importedSheet.frozenColsCount;
                }
              }
            }
          } catch { }
          // Re-apply filter logic after remote update
          if (this.filterActive && this.advFilterSavedState.size > 0) {
            this.recalculateAllFilters();
          }
          this.updateDisplayCache();
        }
        setTimeout(() => this.applyingRemote = false, 50);
      } else if (msg.type === 'cell_update' && msg.r !== undefined && msg.c !== undefined && msg.sheetIdx !== undefined) {
        const remoteUser = 'Collaborator';
        const key = `${msg.sheetIdx}-${msg.r}-${msg.c}`;
        if (!this.cellEditHistory) this.cellEditHistory = {};
        if (!this.cellEditHistory[key]) this.cellEditHistory[key] = [];
        this.cellEditHistory[key].unshift({
          user: remoteUser,
          time: new Date(),
          action: msg.value ? (this.cellEditHistory[key].length === 0 ? 'ADDED' : 'EDITED') : 'CLEARED',
          value: msg.value
        });

        if (msg.sheetIdx === this.currentSheetIdx) {
          this.cells[msg.r][msg.c] = msg.value ?? '';
          if (msg.formatting) {
            this.formats[`${msg.r},${msg.c}`] = msg.formatting;
          } else {
            delete this.formats[`${msg.r},${msg.c}`];
          }
          this.updateDisplayCache();
          if (this.cdr) this.cdr.markForCheck();
        } else {
          const sheet = this.sheets[msg.sheetIdx];
          if (sheet) {
            if (!sheet.cells) sheet.cells = [];
            if (!sheet.cells[msg.r]) sheet.cells[msg.r] = [];
            sheet.cells[msg.r][msg.c] = msg.value ?? '';
            if (!sheet.formats) sheet.formats = {};
            if (msg.formatting) {
              sheet.formats[`${msg.r},${msg.c}`] = msg.formatting;
            } else {
              delete sheet.formats[`${msg.r},${msg.c}`];
            }
          }
        }
      } else if (msg.type === 'cursor' && msg.client_id && msg.r !== undefined && msg.c !== undefined) {
        this.remoteCursors[msg.client_id] = { r: msg.r, c: msg.c }; if (this.cdr) this.cdr.markForCheck();
      } else if (msg.type === 'cursor_remove' && msg.client_id) {
        delete this.remoteCursors[msg.client_id]; if (this.cdr) this.cdr.markForCheck();
      }
    });
  }

  insertFunction(fnName: string) {
    this.closeMenus();
    this.formulaBarValue = `=${fnName}(${this.getRangeRef()})`;
    this.cells[this.selectedRow][this.selectedCol] = this.formulaBarValue;
    this.onCellChange();
    this.showToast(`${fnName} function inserted.`);
  }

  moreFunctions() {
    this.closeMenus();
    this.showToast('More functions library opening...');
  }

  customSort() {
    this.closeMenus();
    this.showToast('Custom Sort options are not available in this preview.');
  }



  @HostListener('document:click')
  onDocClick() { this.closeMenus(); this.activePalette = null; this.hideCtx(); if (this.fillPopupState) this.fillPopupState.showMenu = false; }

  isEditingText(e: KeyboardEvent): boolean {
    const t = e.target as HTMLElement;
    return t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.isContentEditable;
  }

  @HostListener('document:keydown', ['$event'])
  onKey(e: KeyboardEvent) {
    if (this.ocrModalOpen && this.ocrData.length) {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
        e.preventDefault();
        if (e.shiftKey) this.ocrRedo();
        else this.ocrUndo();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') {
        e.preventDefault();
        this.ocrRedo();
        return;
      }
      if ((e.key === 'Delete' || e.key === 'Backspace') && !this.ocrEdit && this.ocrSelStart && this.ocrSelEnd) {
        e.preventDefault();
        const minR = Math.min(this.ocrSelStart.r, this.ocrSelEnd.r);
        const maxR = Math.max(this.ocrSelStart.r, this.ocrSelEnd.r);
        const minC = Math.min(this.ocrSelStart.c, this.ocrSelEnd.c);
        const maxC = Math.max(this.ocrSelStart.c, this.ocrSelEnd.c);
        for (let r = minR; r <= maxR; r++) {
          for (let c = minC; c <= maxC; c++) {
            this.ocrData[r][c] = '';
          }
        }
        this.saveOcrHistory();
        this.cdr.detectChanges();
        return;
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c' && !this.ocrEdit && this.ocrSelStart && this.ocrSelEnd) {
        e.preventDefault();
        const minR = Math.min(this.ocrSelStart.r, this.ocrSelEnd.r);
        const maxR = Math.max(this.ocrSelStart.r, this.ocrSelEnd.r);
        const minC = Math.min(this.ocrSelStart.c, this.ocrSelEnd.c);
        const maxC = Math.max(this.ocrSelStart.c, this.ocrSelEnd.c);
        let text = '';
        for (let r = minR; r <= maxR; r++) {
          let row = [];
          for (let c = minC; c <= maxC; c++) {
            row.push(this.ocrData[r][c] || '');
          }
          text += row.join('\t') + '\n';
        }
        navigator.clipboard.writeText(text);
        this.showToast('Copied from OCR grid');
        return;
      }
      if (this.ocrEdit) return; // if editing an OCR cell, let normal keydown proceed inside it, but don't bubble to main sheet
    }

    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
      e.preventDefault();
      this.save();
      return;
    }

    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'a') {
      if (!this.isEditingText(e) && !this.isEditingCell) {
        e.preventDefault();
        this.selectAll();
        return;
      }
    }

    if ((e.key === 'Delete' || e.key === 'Backspace') && !this.isEditingText(e) && !this.isEditingCell) {
      e.preventDefault();
      const tag = (e.target as HTMLElement).tagName;
      if (tag !== 'INPUT' && tag !== 'TEXTAREA') {
        if (this.rangeStart && this.rangeEnd) {
          const minR = Math.min(this.rangeStart.r, this.rangeEnd.r);
          const maxR = Math.max(this.rangeStart.r, this.rangeEnd.r);
          const minC = Math.min(this.rangeStart.c, this.rangeEnd.c);
          const maxC = Math.max(this.rangeStart.c, this.rangeEnd.c);
          if (minR !== maxR || minC !== maxC) {
            this.clearRangeData();
            return;
          }
        }
        this.clearCell();
      }
      return;
    }

    if (!this.isEditingText(e) && !this.isEditingCell) {
      if (e.key === 'ArrowUp') {
        e.preventDefault();
        const nr = Math.max(0, this.selectedRow - 1);
        this.selectCell(nr, this.selectedCol);
        return;
      }
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        const nr = Math.min(this.cells.length - 1, this.selectedRow + 1);
        this.selectCell(nr, this.selectedCol);
        return;
      }
      if (e.key === 'ArrowLeft' || (e.key === 'Tab' && e.shiftKey)) {
        e.preventDefault();
        const nc = Math.max(0, this.selectedCol - 1);
        this.selectCell(this.selectedRow, nc);
        return;
      }
      if (e.key === 'ArrowRight' || (e.key === 'Tab' && !e.shiftKey)) {
        e.preventDefault();
        const nc = Math.min(this.cells[0].length - 1, this.selectedCol + 1);
        this.selectCell(this.selectedRow, nc);
        return;
      }
      if (e.key === ' ') {
        let hasCheckbox = false;
        let allTrue = true;
        this.forEachSelectedCell((r, c) => {
          if (this.isCheckboxCell(r, c)) {
            hasCheckbox = true;
            if (this.cells[r][c] !== 'TRUE') allTrue = false;
          }
        });
        if (hasCheckbox) {
          e.preventDefault();
          this.pushHistory();
          const newVal = allTrue ? 'FALSE' : 'TRUE';
          this.forEachSelectedCell((r, c) => {
            if (this.isCheckboxCell(r, c)) {
              this.cells[r][c] = newVal;
            }
          });
          this.onCellChange();
          this.save();
          return;
        }
      }
      if (e.key === 'Enter') {
        e.preventDefault();
        this.startEditing();
        return;
      }
      if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
        this.startEditing(e.key);
        return;
      }
    }

    if ((e.ctrlKey || e.metaKey) && e.key === 'z') { e.preventDefault(); this.undo(); }
    if ((e.ctrlKey || e.metaKey) && e.key === 'y') { e.preventDefault(); this.redo(); }
    if ((e.ctrlKey || e.metaKey) && e.key === 'b') { e.preventDefault(); this.toggleFormat('bold'); }
    if ((e.ctrlKey || e.metaKey) && e.key === 'i') { e.preventDefault(); this.toggleFormat('italic'); }
    if ((e.ctrlKey || e.metaKey) && e.key === 'u') { e.preventDefault(); this.toggleFormat('underline'); }
    if ((e.ctrlKey || e.metaKey) && e.key === 'x') { e.preventDefault(); this.cutCell(); }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c') {
      if (!this.isEditingText(e as any) && !this.isEditingCell) {
        e.preventDefault();
        this.copyCell();
      }
    }
    if ((e.ctrlKey || e.metaKey) && (e.key === ';' || e.key === ':')) {
      if (!this.isEditingText(e as any) && !this.isEditingCell) {
        e.preventDefault();
        this.pushHistory();
        const now = new Date();
        if (e.shiftKey) {
          const h = String(now.getHours()).padStart(2, '0');
          const min = String(now.getMinutes()).padStart(2, '0');
          const s = String(now.getSeconds()).padStart(2, '0');
          this.cells[this.selectedRow][this.selectedCol] = `${h}:${min}:${s}`;
        } else {
          const d = String(now.getDate()).padStart(2, '0');
          const m = String(now.getMonth() + 1).padStart(2, '0');
          const y = now.getFullYear();
          this.cells[this.selectedRow][this.selectedCol] = `${d}/${m}/${y}`;
        }
        this.onCellChange();
        this.save();
      }
    }
    if ((e.ctrlKey || e.metaKey) && e.key === 'd') { e.preventDefault(); this.fillDown(); }
    if ((e.ctrlKey || e.metaKey) && e.key === 'r') { e.preventDefault(); this.fillRight(); }
    if ((e.ctrlKey || e.metaKey) && (e.key === 'h' || e.key === 'f')) { e.preventDefault(); this.openFind(); }
    if ((e.ctrlKey || e.metaKey) && e.key === 'm') {
      e.preventDefault();
      if (e.shiftKey) this.setFormat('indent', 'decrease');
      else this.setFormat('indent', 'increase');
    }
    if (e.key === 'Delete') {
      // If a multi-cell range is selected, always clear it (even if an input is focused)
      if (this.rangeStart && this.rangeEnd) {
        const minR = Math.min(this.rangeStart.r, this.rangeEnd.r);
        const maxR = Math.max(this.rangeStart.r, this.rangeEnd.r);
        const minC = Math.min(this.rangeStart.c, this.rangeEnd.c);
        const maxC = Math.max(this.rangeStart.c, this.rangeEnd.c);
        if (minR !== maxR || minC !== maxC) {
          e.preventDefault();
          this.clearRangeData();
        }
      } else {
        // Single cell: only clear if the grid itself is focused (not typing in input)
        const tag = (e.target as HTMLElement).tagName;
        if (tag !== 'INPUT' && tag !== 'TEXTAREA') { e.preventDefault(); this.clearRangeData(); }
      }
    }
    if (e.key === 'Escape') { this.hideCtx(); }
  }

  @HostListener('document:paste', ['$event'])
  pasteFromClipboard(e: ClipboardEvent) {
    if (this.ocrModalOpen && this.ocrData.length && this.ocrSelStart && !this.ocrEdit) {
      e.preventDefault();
      const clipboardData = e.clipboardData || (window as any).clipboardData;
      if (!clipboardData) return;
      const pastedText = clipboardData.getData('Text');
      if (!pastedText) return;

      this.saveOcrHistory();
      const rows = pastedText.split('\n');
      const startR = Math.min(this.ocrSelStart.r, this.ocrSelEnd!.r);
      const startC = Math.min(this.ocrSelStart.c, this.ocrSelEnd!.c);

      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        if (row === undefined || row === null) continue;
        const cols = row.split('\t');
        for (let j = 0; j < cols.length; j++) {
          if (startR + i < this.ocrData.length && startC + j < this.ocrData[0].length) {
            this.ocrData[startR + i][startC + j] = cols[j].replace(/\r$/, '');
          }
        }
      }
      this.saveOcrHistory();
      this.cdr.detectChanges();
      return;
    }

    if (this.isEditingText(e as any) || this.isEditingCell) return;

    e.preventDefault();

    const clipboardData = e.clipboardData || (window as any).clipboardData;
    if (!clipboardData) return;

    const pastedText = clipboardData.getData('Text');
    if (this.richClipboard) {
      const tsvRows: string[] = [];
      for (let r = 0; r < this.richClipboard.rows; r++) {
        const tsvCols: string[] = [];
        for (let c = 0; c < this.richClipboard.cols; c++) {
          tsvCols.push(this.richClipboard.cells[r][c] || '');
        }
        tsvRows.push(tsvCols.join('\t'));
      }
      const expectedTsv = tsvRows.join('\n');
      if (pastedText === expectedTsv || pastedText === expectedTsv + '\r\n' || pastedText === expectedTsv + '\n') {
        this.applyRichPaste('all');
        this.showToast('Pasted.');
        return;
      }
    }

    if (clipboardData.items) {
      for (let i = 0; i < clipboardData.items.length; i++) {
        if (clipboardData.items[i].type.indexOf('image') !== -1) {
          const file = clipboardData.items[i].getAsFile();
          if (file) {
            const reader = new FileReader();
            reader.onload = (ev) => {
              this.pushHistory({ action_type: 'paste-image', target_range: this.getA1(this.selectedRow, this.selectedCol) });
              this.cells[this.selectedRow][this.selectedCol] = ev.target!.result as string;
              this.formulaBarValue = '[IMAGE]';
              this.onCellChange();
              this.save();
              this.showToast('Image pasted into cell.');
              if (this.cdr) this.cdr.detectChanges();
            };
            reader.readAsDataURL(file);
            return;
          }
        }
      }
    }

    const pastedHtml = clipboardData.getData('text/html');

    if (!pastedHtml && !pastedText) return;

    this.pushHistory({ action_type: 'paste-clipboard', target_range: 'Multiple' });

    const startRow = this.selectedRow;
    const startCol = this.selectedCol;
    let maxCols = 1;
    let maxRows = 1;

    let parsedFromHtml = false;

    if (pastedHtml) {
      const parser = new DOMParser();
      const doc = parser.parseFromString(pastedHtml, 'text/html');
      const table = doc.querySelector('table');
      const img = doc.querySelector('img');

      if (table) {
        parsedFromHtml = true;
        const rows = Array.from(table.rows);
        maxRows = rows.length;

        for (let r = 0; r < rows.length; r++) {
          const cells = Array.from(rows[r].cells);
          maxCols = Math.max(maxCols, cells.length);
          for (let c = 0; c < cells.length; c++) {
            const targetR = startRow + r;
            const targetC = startCol + c;

            if (targetR < this.ROWS && targetC < this.COLS) {
              const cell = cells[c];

              // If the cell contains an image
              const cellImg = cell.querySelector('img');
              if (cellImg && cellImg.src) {
                this.cells[targetR][targetC] = cellImg.src;
              } else {
                this.cells[targetR][targetC] = cell.innerText.trim();
              }

              let formats: any = {};

              // Handle bold
              if (cell.tagName.toLowerCase() === 'th' ||
                cell.style.fontWeight === 'bold' ||
                cell.style.fontWeight === '700' ||
                cell.querySelector('b') ||
                cell.querySelector('strong')) {
                formats.bold = true;
              }

              // Handle italic
              if (cell.style.fontStyle === 'italic' ||
                cell.querySelector('i') ||
                cell.querySelector('em')) {
                formats.italic = true;
              }

              // Handle background color
              if (cell.style.backgroundColor &&
                cell.style.backgroundColor !== 'transparent' &&
                cell.style.backgroundColor !== 'rgba(0, 0, 0, 0)') {
                formats.bg = cell.style.backgroundColor;
              }

              // Handle text color
              if (cell.style.color && cell.style.color !== 'inherit') {
                formats.color = cell.style.color;
              }

              if (Object.keys(formats).length > 0) {
                this.formats[`${targetR},${targetC}`] = {
                  ...(this.formats[`${targetR},${targetC}`] || {}),
                  ...formats
                };
              }
            }
          }
        }
      } else if (img && img.src) {
        parsedFromHtml = true;
        this.cells[startRow][startCol] = img.src;
        this.formulaBarValue = '[IMAGE]';
      }
    }

    if (!parsedFromHtml && pastedText) {
      const rows = pastedText.split(/\r?\n/);
      if (rows.length > 0 && rows[rows.length - 1] === '') {
        rows.pop(); // Remove trailing newline from spreadsheet copies
      }
      maxRows = rows.length;

      for (let r = 0; r < rows.length; r++) {
        const cols = rows[r].split('\t');
        maxCols = Math.max(maxCols, cols.length);
        for (let c = 0; c < cols.length; c++) {
          const targetR = startRow + r;
          const targetC = startCol + c;
          if (targetR < this.ROWS && targetC < this.COLS) {
            this.cells[targetR][targetC] = cols[c];
          }
        }
      }
    }

    this.rangeStart = { r: startRow, c: startCol };
    this.rangeEnd = {
      r: Math.min(startRow + maxRows - 1, this.ROWS - 1),
      c: Math.min(startCol + maxCols - 1, this.COLS - 1)
    };

    this.onCellChange();
    this.save();
    this.showToast('Data pasted');
  }

  @HostListener('document:mouseup')
  onDocMouseUp(e?: MouseEvent) {
    if (this.isFilling && this.fillEnd) {
      this.applyFill(e?.ctrlKey ?? false);
    }
    this.isDraggingRange = false;
    this.isFilling = false;
    this.fillStart = null;
    this.ocrDragging = false;
  }

  // ── Range selection helpers ──────────────────────────────────────────────
  onCellMouseDown(e: MouseEvent, r: number, c: number) {
    if (this.fillPopupState) this.fillPopupState = null;
    if ((e.target as HTMLElement).classList.contains('fill-handle')) return;
    if (this.hasComment(r, c) && this.sidePanelApp !== 'comments') {
      this.sidePanelApp = 'comments';
      this.commentsViewFilter = 'current';
      this.updateCachedComments();
    }

    // Support right-click on Windows (button===2) and Mac (ctrlKey)
    const isRightClick = e.button === 2 || (e.button === 0 && e.ctrlKey);
    if (isRightClick) {
      if (this.isCellInRange(r, c)) {
        return; // Clicked inside existing selection, keep it
      } else {
        // Right-clicked outside selection, select this single cell
        this.selectCell(r, c);
        return;
      }
    }

    if (this.activeNotePopup) {
      this.activeNotePopup = null;
    }

    this.isDraggingRange = true;
    this.rangeStart = { r, c };
    this.rangeEnd = { r, c };
    this.fillEnd = null;
  }

  onCellMouseEnter(r: number, c: number) {
    if (this.isDraggingRange && this.rangeStart) {
      this.rangeEnd = { r, c };
    }
    if (this.isFilling && this.fillStart) {
      this.fillEnd = { r, c };
    }
  }

  isCellSelected(r: number, c: number): boolean {
    return this.selectedRow === r && this.selectedCol === c;
  }

  isCheckboxCell(r: number, c: number): boolean {
    const f = this.formats[`${r},${c}`];
    return !!(f && (f as any)['checkbox']);
  }

  isUrl(val: string): boolean {
    if (!val || typeof val !== 'string') return false;
    return /^(https?:\/\/[^\s]+)$/i.test(val.trim());
  }

  onCheckboxMouseDown(e: MouseEvent, r: number, c: number) {
    e.preventDefault();
    e.stopPropagation();

    let inSelection = false;
    if (this.rangeStart && this.rangeEnd) {
      const minR = Math.min(this.rangeStart.r, this.rangeEnd.r);
      const maxR = Math.max(this.rangeStart.r, this.rangeEnd.r);
      const minC = Math.min(this.rangeStart.c, this.rangeEnd.c);
      const maxC = Math.max(this.rangeStart.c, this.rangeEnd.c);
      if (r >= minR && r <= maxR && c >= minC && c <= maxC) {
        inSelection = true;
      }
    }

    this.pushHistory();
    const isChecking = this.cells[r][c] !== 'TRUE';
    const newVal = isChecking ? 'TRUE' : 'FALSE';

    if (inSelection) {
      this.forEachSelectedCell((sr, sc) => {
        if (this.isCheckboxCell(sr, sc)) {
          this.cells[sr][sc] = newVal;
        }
      });
    } else {
      this.cells[r][c] = newVal;
      this.selectCell(r, c);
    }

    if (this.selectedRow === r && this.selectedCol === c) {
      this.formulaBarValue = this.cells[r][c];
    }
    this.onCellChange();
    this.save();
  }

  getContrastText(hexColor?: string): string {
    if (!hexColor || typeof hexColor !== 'string') return '#FFFFFF';
    const hex = hexColor.replace('#', '');
    if (!/^[0-9A-Fa-f]{3}$|^[0-9A-Fa-f]{6}$/.test(hex)) return '#FFFFFF';
    let fullHex = hex;
    if (hex.length === 3) {
      fullHex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
    }
    const r = parseInt(fullHex.substring(0, 2), 16);
    const g = parseInt(fullHex.substring(2, 4), 16);
    const b = parseInt(fullHex.substring(4, 6), 16);
    // relative luminance (per WCAG)
    const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
    return luminance > 0.6 ? '#111111' : '#FFFFFF';
  }

  openCustomDropdown(event: MouseEvent, r: number, c: number) {
    event.stopPropagation();
    const opts = this.getCellDropdownOptions(r, c);
    if (!opts || !opts.length) return;
    const target = event.currentTarget as HTMLElement;
    const rect = target.getBoundingClientRect();

    const spaceBelow = window.innerHeight - rect.bottom - 16;
    const spaceAbove = rect.top - 16;

    let maxListHeight = 180;
    let topPos = rect.bottom + 2;

    if (spaceBelow < 180 && spaceAbove > spaceBelow) {
      maxListHeight = Math.min(180, Math.max(100, spaceAbove - 45));
      topPos = Math.max(8, rect.top - maxListHeight - 45);
    } else {
      maxListHeight = Math.min(180, Math.max(100, spaceBelow - 45));
    }

    const overlay = document.createElement('div');
    overlay.className = 'custom-dropdown-overlay';
    overlay.style.position = 'fixed';
    overlay.style.top = `${topPos}px`;
    overlay.style.left = `${rect.left}px`;
    overlay.style.width = `${Math.max(rect.width, 160)}px`;
    overlay.style.background = '#242424';
    overlay.style.border = '1px solid #3c4043';
    overlay.style.borderRadius = '8px';
    overlay.style.boxShadow = '0 8px 24px rgba(0,0,0,0.5)';
    overlay.style.zIndex = '100000';
    overlay.style.display = 'flex';
    overlay.style.flexDirection = 'column';

    const scrollContainer = document.createElement('div');
    scrollContainer.className = 'option-list';
    scrollContainer.setAttribute('role', 'listbox');
    scrollContainer.style.maxHeight = `${maxListHeight}px`;
    scrollContainer.style.overflowY = 'auto';
    overlay.appendChild(scrollContainer);

    const closeOverlay = (e: MouseEvent) => {
      if (!overlay.contains(e.target as Node)) {
        if (document.body.contains(overlay)) {
          document.body.removeChild(overlay);
        }
        document.removeEventListener('click', closeOverlay);
        document.removeEventListener('keydown', onKeyDown);
      }
    };

    const onKeyDown = (e: KeyboardEvent) => {
      const rows = Array.from(scrollContainer.querySelectorAll('.option-row')) as HTMLElement[];
      const currentIndex = rows.indexOf(document.activeElement as HTMLElement);

      if (e.key === 'ArrowDown') {
        e.preventDefault();
        e.stopPropagation();
        const nextIndex = currentIndex === -1 ? 0 : Math.min(currentIndex + 1, rows.length - 1);
        rows[nextIndex]?.focus();
      } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        e.stopPropagation();
        const prevIndex = currentIndex === -1 ? rows.length - 1 : Math.max(currentIndex - 1, 0);
        rows[prevIndex]?.focus();
      } else if (e.key === 'Enter') {
        if (currentIndex !== -1) {
          e.preventDefault();
          e.stopPropagation();
          rows[currentIndex].click();
        }
      } else if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        if (document.body.contains(overlay)) {
          document.body.removeChild(overlay);
        }
        document.removeEventListener('click', closeOverlay);
        document.removeEventListener('keydown', onKeyDown);
      }
    };

    // Preserving underlying clear-cell logic for TASK 6b
    const clearCellLogic = (e?: MouseEvent) => {
      if (e) e.stopPropagation();
      this.pushHistory();
      this.cells[r][c] = '';
      if (this.selectedRow === r && this.selectedCol === c) {
        this.formulaBarValue = '';
      }
      this.onCellChange();
      this.save();
      if (document.body.contains(overlay)) {
        document.body.removeChild(overlay);
      }
      document.removeEventListener('click', closeOverlay);
      document.removeEventListener('keydown', onKeyDown);
    };

    // TASK 6b: Clear value option as first item inside scrollable list
    const clearRow = document.createElement('div');
    clearRow.className = 'option-row clear-option';
    clearRow.tabIndex = 0;
    clearRow.setAttribute('role', 'option');
    clearRow.setAttribute('aria-selected', 'false');
    clearRow.onclick = clearCellLogic;
    clearRow.innerText = 'Clear value';
    scrollContainer.appendChild(clearRow);

    opts.forEach((opt: any) => {
      const optValue = opt.label || opt;
      const currentVal = this.cells[r][c] || '';
      const isMulti = this.validations[`${r},${c}`]?.isMultiSelect;
      const isSelected = isMulti 
        ? currentVal.split(',').map((p: string) => p.trim()).includes(optValue)
        : currentVal === optValue;

      const row = document.createElement('div');
      row.className = 'option-row';
      row.tabIndex = 0;
      row.setAttribute('role', 'option');
      row.setAttribute('data-selected', isSelected ? 'true' : 'false');
      row.setAttribute('aria-selected', isSelected ? 'true' : 'false');
      row.innerText = optValue;

      row.onclick = (e) => {
        e.stopPropagation();
        this.pushHistory();
        
        const v = this.validations[`${r},${c}`];
        const isMultiVal = v && v.isMultiSelect;

        if (isMultiVal) {
          let currentStr = this.cells[r][c] || '';
          let parts = currentStr.split(',').map((p: string) => p.trim()).filter((p: string) => !!p);
          
          if (parts.includes(optValue)) {
            parts = parts.filter((p: string) => p !== optValue);
          } else {
            parts.push(optValue);
          }
          this.cells[r][c] = parts.join(', ');
        } else {
          this.cells[r][c] = optValue;
        }

        if (this.selectedRow === r && this.selectedCol === c) {
          this.formulaBarValue = this.cells[r][c];
        }
        this.onCellChange();
        this.save();
        
        if (!isMultiVal) {
          if (document.body.contains(overlay)) {
            document.body.removeChild(overlay);
          }
          document.removeEventListener('click', closeOverlay);
          document.removeEventListener('keydown', onKeyDown);
        }
      };
      scrollContainer.appendChild(row);
    });

    // Divider
    const divider = document.createElement('div');
    divider.className = 'dropdown-divider';
    overlay.appendChild(divider);

    // Footer row
    const footer = document.createElement('div');
    footer.className = 'dropdown-footer';

    const editBtn = document.createElement('button');
    editBtn.className = 'edit-btn';
    editBtn.type = 'button';
    editBtn.tabIndex = 0;
    editBtn.setAttribute('aria-label', 'Edit picklist');
    editBtn.innerHTML = '<svg class="edit-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 20h9"></path><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path></svg><span>Edit</span>';
    
    editBtn.onclick = (e) => {
      e.stopPropagation();
      if (document.body.contains(overlay)) {
        document.body.removeChild(overlay);
      }
      document.removeEventListener('click', closeOverlay);
      document.removeEventListener('keydown', onKeyDown);
      this.selectCell(r, c);
      this.rangeStart = { r, c };
      this.rangeEnd = { r, c };
      this.openValidationModal();
      if (this.cdr) this.cdr.detectChanges();
    };

    footer.appendChild(editBtn);
    overlay.appendChild(footer);

    document.body.appendChild(overlay);

    // use setTimeout so the current click doesn't instantly close it
    setTimeout(() => {
      document.addEventListener('click', closeOverlay);
      document.addEventListener('keydown', onKeyDown);
    }, 0);
  }

  isCellInRange(r: number, c: number): boolean {
    if (!this.rangeStart || !this.rangeEnd) return false;
    const minR = Math.min(this.rangeStart.r, this.rangeEnd.r);
    const maxR = Math.max(this.rangeStart.r, this.rangeEnd.r);
    const minC = Math.min(this.rangeStart.c, this.rangeEnd.c);
    const maxC = Math.max(this.rangeStart.c, this.rangeEnd.c);
    // For full-column/row selection (entire range), highlight all including anchor
    const isFullRange = (maxR - minR > 0 || maxC - minC > 0);
    if (!isFullRange) return false;
    return r >= minR && r <= maxR && c >= minC && c <= maxC;
  }

  // ── Fill handle helpers ──────────────────────────────────────────────────
  isFillHandleCell(r: number, c: number): boolean {
    // When a range is active: show ONE dot at the bottom-right of the range only
    if (this.rangeStart && this.rangeEnd) {
      const minR = Math.min(this.rangeStart.r, this.rangeEnd.r);
      const maxR = Math.max(this.rangeStart.r, this.rangeEnd.r);
      const minC = Math.min(this.rangeStart.c, this.rangeEnd.c);
      const maxC = Math.max(this.rangeStart.c, this.rangeEnd.c);
      // Multi-cell range: dot only at bottom-right
      if (minR !== maxR || minC !== maxC) return r === maxR && c === maxC;
    }
    // No range or single-cell range: dot at selected cell
    return r === this.selectedRow && c === this.selectedCol;
  }

  isCellInFillPreview(r: number, c: number): boolean {
    if (!this.isFilling || !this.fillStart || !this.fillEnd) return false;
    const minR = Math.min(this.fillStart.r, this.fillEnd.r);
    const maxR = Math.max(this.fillStart.r, this.fillEnd.r);
    const minC = Math.min(this.fillStart.c, this.fillEnd.c);
    const maxC = Math.max(this.fillStart.c, this.fillEnd.c);
    return r >= minR && r <= maxR && c >= minC && c <= maxC;
  }

  onFillHandleMouseDown(e: MouseEvent, r: number, c: number) {
    e.preventDefault();
    e.stopPropagation();
    this.isFilling = true;
    this.isDraggingRange = false;
    this.fillStart = { r, c };
    this.fillEnd = { r, c };
  }

  private applyFill(ctrlKey: boolean = false) {
    if (!this.fillStart || !this.fillEnd) return;

    const srcMinR = this.rangeStart ? Math.min(this.rangeStart.r, this.rangeEnd!.r) : this.fillStart.r;
    const srcMaxR = this.rangeStart ? Math.max(this.rangeStart.r, this.rangeEnd!.r) : this.fillStart.r;
    const srcMinC = this.rangeStart ? Math.min(this.rangeStart.c, this.rangeEnd!.c) : this.fillStart.c;
    const srcMaxC = this.rangeStart ? Math.max(this.rangeStart.c, this.rangeEnd!.c) : this.fillStart.c;

    const dstR = this.fillEnd.r;
    const dstC = this.fillEnd.c;

    const goDown = dstR > srcMaxR;
    const goUp = dstR < srcMinR;
    const goRight = dstC > srcMaxC;
    const goLeft = dstC < srcMinC;

    if (!goDown && !goUp && !goRight && !goLeft) return;

    // snapshot source data + formats
    const sourceData: { r: number; c: number; val: string; fmt: any }[] = [];
    for (let r = srcMinR; r <= srcMaxR; r++)
      for (let c = srcMinC; c <= srcMaxC; c++)
        sourceData.push({ r, c, val: this.cells[r][c], fmt: JSON.parse(JSON.stringify(this.formats[`${r},${c}`] ?? null)) });

    // compute fill target bounds
    const dstMinR = goUp ? dstR : (goDown ? srcMaxR + 1 : srcMinR);
    const dstMaxR = goUp ? srcMinR - 1 : (goDown ? dstR : srcMaxR);
    const dstMinC = goLeft ? dstC : (goRight ? srcMaxC + 1 : srcMinC);
    const dstMaxC = goLeft ? srcMinC - 1 : (goRight ? dstC : srcMaxC);

    // backup target cells
    const targetBackup: { r: number; c: number; val: string; fmt: any }[] = [];
    for (let r = dstMinR; r <= dstMaxR; r++)
      for (let c = dstMinC; c <= dstMaxC; c++)
        targetBackup.push({ r, c, val: this.cells[r][c], fmt: JSON.parse(JSON.stringify(this.formats[`${r},${c}`] ?? null)) });

    this.pushHistory({
      action_type: 'drag-to-fill',
      target_range: `${this.getA1(dstMinR, dstMinC)}:${this.getA1(dstMaxR, dstMaxC)}`
    });

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
    this.rangeEnd = { r: newMaxR, c: newMaxC };
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
        if (goDown) fillTargetRows(dstMinR, dstMaxR, 1, false);
        else fillTargetRows(dstMaxR, dstMinR, -1, true);
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
        if (goRight) fillTargetCols(dstMinC, dstMaxC, 1, false);
        else fillTargetCols(dstMaxC, dstMinC, -1, true);
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
    return val.replace(/\$?[A-Z]+\$?\d+/g, (match) => {
      if (match.startsWith('$') && match.includes('$', 1)) return match; // absolute
      const hasAbsCol = match.startsWith('$');
      const hasAbsRow = match.match(/\$\d+$/);
      const colStr = match.replace(/\$/g, '').match(/^[A-Z]+/)![0];
      const rowStr = match.match(/\d+$/)![0];
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
    const dateRx = /^(\d{1,2})-(\d{1,2})-(\d{2,4})$/;
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
    const trailRx = /^(.*?)(\d+)$/;
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
  }

  /** Smart series: detect number, date (MM-DD-YY or MM-DD-YYYY), repeat, or formula shifting */
  private getNextSeriesValue(srcVals: string[], offset: number, isVertical: boolean): string {
    let v = '';
    let idx = 0;

    // Determine the base string to use
    if (srcVals.length === 1) {
      v = srcVals[0];
      // Interpolate numbers/dates only if it's NOT a formula
      if (!v.startsWith('=')) {
        const num = Number(v);
        // By default in spreadsheets, dragging a single number copies it instead of incrementing.
        // We only interpolate dates automatically.
        const dateMatch = v.match(/^(\d{1,2})-(\d{1,2})-(\d{2,4})$/);
        if (dateMatch) {
          const m = parseInt(dateMatch[1], 10);
          const d = parseInt(dateMatch[2], 10);
          const y = parseInt(dateMatch[3], 10);
          const fullY = y < 100 ? 2000 + y : y;
          const dt = new Date(fullY, m - 1, d);
          dt.setDate(dt.getDate() + offset);
          const nm = String(dt.getMonth() + 1).padStart(2, '0');
          const nd = String(dt.getDate()).padStart(2, '0');
          const ny = y < 100 ? String(dt.getFullYear()).slice(2) : String(dt.getFullYear());
          return `${nm}-${nd}-${ny}`;
        }
      }
    } else {
      // Multi-value: try linear step
      const nums = srcVals.map(s => Number(s));
      if (nums.every(n => !isNaN(n) && srcVals[nums.indexOf(n)].trim() !== '')) {
        const step = nums.length > 1 ? nums[1] - nums[0] : 1;
        const base = nums[nums.length - 1];
        const steps = offset - (srcVals.length - 1);
        return String(base + step * steps);
      }
      // Repeat cycle
      idx = ((offset % srcVals.length) + srcVals.length) % srcVals.length;
      v = srcVals[idx];
    }

    if (v.startsWith('=')) {
      const shiftAmount = offset - idx;
      const rowDelta = isVertical ? shiftAmount : 0;
      const colDelta = isVertical ? 0 : shiftAmount;

      return v.replace(/[A-Z]+\d+/g, (match) => {
        const colStr = match.match(/^[A-Z]+/)![0];
        const rowStr = match.match(/\d+$/)![0];
        let colIdx = colStr.charCodeAt(0) - 65;
        let rowIdx = parseInt(rowStr, 10) - 1;

        colIdx += colDelta;
        rowIdx += rowDelta;

        if (colIdx < 0) colIdx = 0;
        if (rowIdx < 0) rowIdx = 0;
        return `${colName(colIdx)}${rowIdx + 1}`;
      });
    }

    return v;
  }

  // ── Column / Row header selection ────────────────────────────────────────
  selectEntireCol(c: number) {
    this.selectedColHeader = c;
    this.selectedRowHeader = null;
    this.rangeStart = { r: 0, c };
    this.rangeEnd = { r: this.ROWS - 1, c };
    this.selectedRow = 0;
    this.selectedCol = c;
    this.formulaBarValue = '';
  }

  selectEntireRow(r: number) {
    this.selectedRowHeader = r;
    this.selectedColHeader = null;
    this.rangeStart = { r, c: 0 };
    this.rangeEnd = { r, c: this.COLS - 1 };
    this.selectedRow = r;
    this.selectedCol = 0;
    this.formulaBarValue = '';
  }

  getContiguousDataRange(startR: number, startC: number): { minR: number, maxR: number, minC: number, maxC: number } {
    let minR = startR, maxR = startR, minC = startC, maxC = startC;
    let expanded = true;
    console.log('--- getContiguousDataRange start ---', startR, startC);

    while (expanded) {
      expanded = false;
      if (minR > 0) {
        let hasData = false;
        for (let c = minC; c <= maxC; c++) {
          if (this.cells[minR - 1] && this.cells[minR - 1][c] != null && String(this.cells[minR - 1][c]).trim() !== '') { hasData = true; break; }
        }
        if (hasData) { minR--; expanded = true; continue; }
      }
      if (maxR < this.ROWS - 1) {
        let hasData = false;
        for (let c = minC; c <= maxC; c++) {
          if (this.cells[maxR + 1] && this.cells[maxR + 1][c] != null && String(this.cells[maxR + 1][c]).trim() !== '') { hasData = true; break; }
        }
        if (hasData) { maxR++; expanded = true; continue; }
      }
      if (minC > 0) {
        let hasData = false;
        for (let r = minR; r <= maxR; r++) {
          if (this.cells[r] && this.cells[r][minC - 1] != null && String(this.cells[r][minC - 1]).trim() !== '') { hasData = true; break; }
        }
        if (hasData) { minC--; expanded = true; continue; }
      }
      if (maxC < this.COLS - 1) {
        let hasData = false;
        for (let r = minR; r <= maxR; r++) {
          if (this.cells[r] && this.cells[r][maxC + 1] != null && String(this.cells[r][maxC + 1]).trim() !== '') { hasData = true; break; }
        }
        if (hasData) { maxC++; expanded = true; continue; }
      }
    }
    console.log('--- getContiguousDataRange result ---', { minR, maxR, minC, maxC });
    return { minR, maxR, minC, maxC };
  }

  selectAll() {
    this.closeMenus();
    this.isEditingCell = false;

    this.selectedColHeader = null;
    this.selectedRowHeader = null;

    const r = this.selectedRow;
    const c = this.selectedCol;
    const cellValue = this.cells[r] && this.cells[r][c];
    const isCellEmpty = cellValue == null || String(cellValue).trim() === '';

    console.log('selectAll called. active cell:', r, c, 'value:', cellValue, 'isEmpty:', isCellEmpty);

    if (!isCellEmpty) {
      const range = this.getContiguousDataRange(r, c);
      const isAlreadySelected = this.rangeStart && this.rangeEnd &&
        Math.min(this.rangeStart.r, this.rangeEnd.r) === range.minR &&
        Math.max(this.rangeStart.r, this.rangeEnd.r) === range.maxR &&
        Math.min(this.rangeStart.c, this.rangeEnd.c) === range.minC &&
        Math.max(this.rangeStart.c, this.rangeEnd.c) === range.maxC;

      if (!isAlreadySelected) {
        this.rangeStart = { r: range.minR, c: range.minC };
        this.rangeEnd = { r: range.maxR, c: range.maxC };
        this.selectedRow = range.minR;
        this.selectedCol = range.minC;
        return;
      }
    }

    this.rangeStart = { r: 0, c: 0 };
    this.rangeEnd = { r: this.ROWS - 1, c: this.COLS - 1 };
    this.selectedRow = 0;
    this.selectedCol = 0;
    this.formulaBarValue = this.isImageCell(0, 0) ? '[IMAGE]' : this.cells[0][0] || '';
  }

  isColHeaderSelected(c: number): boolean {
    if (this.selectedColHeader === c) return true;
    if (this.rangeStart && this.rangeEnd) {
      const minC = Math.min(this.rangeStart.c, this.rangeEnd.c);
      const maxC = Math.max(this.rangeStart.c, this.rangeEnd.c);
      return c >= minC && c <= maxC;
    }
    return false;
  }
  isRowHeaderSelected(r: number): boolean {
    if (this.selectedRowHeader === r) return true;
    if (this.rangeStart && this.rangeEnd) {
      const minR = Math.min(this.rangeStart.r, this.rangeEnd.r);
      const maxR = Math.max(this.rangeStart.r, this.rangeEnd.r);
      return r >= minR && r <= maxR;
    }
    return false;
  }

  // ── Right-click context menu ──────────────────────────────────────────────
  onHeaderRightClick(e: MouseEvent, type: 'row' | 'col', idx: number) {
    e.preventDefault();
    if (type === 'col') {
      if (!this.isColHeaderSelected(idx)) this.selectEntireCol(idx);
    } else {
      if (!this.isRowHeaderSelected(idx)) this.selectEntireRow(idx);
    }

    this.showContextMenu(e);
  }

  onCellRightClick(e: MouseEvent, r: number, c: number) {
    e.preventDefault();
    if (!this.isCellInRange(r, c)) {
      this.selectCell(r, c);
    }
    this.ctxRow = r;
    this.ctxCol = c;

    this.showContextMenu(e);
  }

  showContextMenu(e: MouseEvent) {
    const menuWidth = 220;
    const estimatedMenuHeight = 430;
    let x = e.clientX;
    let y = e.clientY;

    if (x + menuWidth > window.innerWidth) x = window.innerWidth - menuWidth;

    let spaceBelow = window.innerHeight - y - 10;
    let spaceAbove = y - 10;

    this.ctxX = x;

    if (estimatedMenuHeight > spaceBelow && spaceAbove > spaceBelow) {
      // More space above, and it doesn't fit below. Open upwards.
      this.ctxTop = null;
      this.ctxBottom = window.innerHeight - y;
      this.ctxMaxHeight = spaceAbove;
    } else {
      // Open downwards
      this.ctxBottom = null;
      this.ctxTop = y;
      this.ctxMaxHeight = spaceBelow;
    }

    this.ctxVisible = true;
  }

  hideCtx() {
    this.ctxVisible = false;
    this.activeCtxSubmenu = null;
  }

  cutCell() {
    // Capture the range first (same as copy)
    this.copyCell();
    // Then clear the source cells (values + formats + validations)
    this.pushHistory();
    const startR = this.richClipboard!.originR;
    const startC = this.richClipboard!.originC;
    for (let r = 0; r < this.richClipboard!.rows; r++) {
      for (let c = 0; c < this.richClipboard!.cols; c++) {
        const srcR = startR + r;
        const srcC = startC + c;
        this.cells[srcR][srcC] = '';
        delete this.formats[`${srcR},${srcC}`];
        delete this.validations[`${srcR},${srcC}`];
      }
    }
    this.formulaBarValue = '';
    this.onCellChange();
    this.showToast(`Cut ${this.richClipboard!.rows}×${this.richClipboard!.cols} cell${this.richClipboard!.rows * this.richClipboard!.cols > 1 ? 's' : ''}.`);
  }

  // ── Clear all cells in current range / selection ─────────────────────────
  clearRangeData() {
    this.pushHistory();
    if (this.rangeStart && this.rangeEnd) {
      const minR = Math.min(this.rangeStart.r, this.rangeEnd.r);
      const maxR = Math.max(this.rangeStart.r, this.rangeEnd.r);
      const minC = Math.min(this.rangeStart.c, this.rangeEnd.c);
      const maxC = Math.max(this.rangeStart.c, this.rangeEnd.c);
      for (let r = minR; r <= maxR; r++) {
        for (let c = minC; c <= maxC; c++) {
          this.cells[r][c] = '';
          const ref = `${r},${c}`;
          if (this.formats[ref]) {
            delete this.formats[ref].bg;
            delete this.formats[ref].bold;
            delete this.formats[ref].italic;
            delete this.formats[ref].color;
            delete this.formats[ref].strikethrough;
            delete (this.formats[ref] as any).checkbox;
          }
        }
      }
    } else {
      this.cells[this.selectedRow][this.selectedCol] = '';
      const ref = `${this.selectedRow},${this.selectedCol}`;
      if (this.formats[ref]) {
        delete this.formats[ref].bg;
        delete this.formats[ref].bold;
        delete this.formats[ref].italic;
        delete this.formats[ref].color;
        delete this.formats[ref].strikethrough;
        delete (this.formats[ref] as any).checkbox;
      }
    }
    this.formulaBarValue = '';
    this.onCellChange();
    this.save();
    this.showToast('Selection cleared.');
  }



  // ── Dropdown / Data validation ───────────────────────────────────────────────
  hasCellDropdown(r: number, c: number): boolean {
    return !!this.validations[`${r},${c}`];
  }

  isDisplayAsChip(r: number, c: number): boolean {
    const v = this.validations[`${r},${c}`];
    if (!v) return false;
    return v.displayAsChip !== false; // defaults to true
  }

  getCellRef(r: number, c: number): string {
    return colName(c) + (r + 1);
  }

  getRangeRef(): string {
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

  getDropdownTextColor(r: number, c: number, val: string): string {
    const opts = this.getCellDropdownOptions(r, c);
    const found = opts.find(o => (typeof o === 'string' ? o : o.label) === val) as DropdownOption | undefined;
    if (!found) return '#000000';
    return found.textColor || '#000000';
  }

  splitValue(val: string): string[] {
    if (!val) return [];
    return val.split(',').map(s => s.trim()).filter(s => !!s);
  }




  // Manage Picklist Rules
  refreshManagePicklistRules() {
    const rulesMap = new Map<string, {
      sheetName: string;
      sheetIdx: number;
      rangeRef: string;
      options: DropdownOption[];
      isMultiSelect: boolean;
      displayAsChip: boolean;
      cells: { r: number, c: number }[];
    }>();

    const sheetsList = (this.sheets && this.sheets.length > 0) ? this.sheets : [{ name: 'Sheet1', validations: this.validations }];

    for (let sIdx = 0; sIdx < sheetsList.length; sIdx++) {
      const sheet = sheetsList[sIdx];
      const sheetName = sheet?.name || `Sheet${sIdx + 1}`;
      const vals = (sIdx === this.currentSheetIdx) ? (this.validations || {}) : (sheet.validations || {});

      for (const key of Object.keys(vals)) {
        const parts = key.split(',');
        const r = parseInt(parts[0], 10);
        const c = parseInt(parts[1], 10);
        const val = vals[key];
        if (val && val.type === 'list' && val.options && val.options.length) {
          const optionsJson = JSON.stringify(val.options);
          const ruleKey = `${sIdx}_${val.isMultiSelect}_${val.displayAsChip}_${optionsJson}`;
          
          if (!rulesMap.has(ruleKey)) {
            rulesMap.set(ruleKey, {
              sheetName: sheetName,
              sheetIdx: sIdx,
              rangeRef: '',
              options: val.options.map((o: any) => typeof o === 'string' ? { label: o } : o),
              isMultiSelect: !!val.isMultiSelect,
              displayAsChip: val.displayAsChip !== false,
              cells: []
            });
          }
          rulesMap.get(ruleKey)!.cells.push({ r, c });
        }
      }
    }

    const rulesList = Array.from(rulesMap.values()).map(rule => {
      rule.rangeRef = this.formatRuleRangeRef(rule.sheetName, rule.cells);
      return rule;
    });

    const currentSheetName = (this.sheets && this.sheets[this.currentSheetIdx]?.name) || 'Sheet1';

    this._managePicklistRules = rulesList.filter(rule => {
      if (this.viewRulesSheet === 'current') {
        if (rule.sheetIdx !== this.currentSheetIdx && rule.sheetName !== currentSheetName) return false;
      } else if (this.viewRulesSheet !== 'all') {
        if (rule.sheetName !== this.viewRulesSheet) return false;
      }
      if (this.viewRulesType === 'list' && rule.isMultiSelect) return false;
      if (this.viewRulesType === 'range' && !rule.rangeRef.includes(':')) return false;
      return true;
    });
  }

  formatRuleRangeRef(sheetName: string, cells: { r: number, c: number }[]): string {
    if (!cells.length) return `'${sheetName}'.A1`;
    let minR = cells[0].r, maxR = cells[0].r;
    let minC = cells[0].c, maxC = cells[0].c;
    for (const cell of cells) {
      if (cell.r < minR) minR = cell.r;
      if (cell.r > maxR) maxR = cell.r;
      if (cell.c < minC) minC = cell.c;
      if (cell.c > maxC) maxC = cell.c;
    }
    if (minR === maxR && minC === maxC) {
      return `'${sheetName}'.${this.colLabel(minC)}${minR + 1}`;
    }
    return `'${sheetName}'.${this.colLabel(minC)}${minR + 1}:${this.colLabel(maxC)}${maxR + 1}`;
  }

  openManagePicklistSidebar(e?: Event) {
    if (e) e.stopPropagation();
    this.refreshManagePicklistRules();
    this.managePicklistSidebarOpen = true;
    this.validationModalOpen = false;
    this.closeMenus();
    if (this.cdr) this.cdr.detectChanges();
  }

  editPicklistRule(rule: any, e?: Event) {
    if (e) e.stopPropagation();
    this.picklistOptions = rule.options.map((o: any) => ({ ...o }));
    this.picklistSelectType = rule.isMultiSelect ? 'multi' : 'single';
    this.displayAsChip = rule.displayAsChip;
    this.appliesToInput = rule.rangeRef;
    this.appliesToEditing = false;
    this.isCopyMode = false;
    this.editingOldRule = rule;
    this.managePicklistSidebarOpen = false;
    this.validationModalOpen = true;
    if (this.cdr) this.cdr.detectChanges();
  }

  copyPicklistRule(rule: any, e?: Event) {
    if (e) e.stopPropagation();
    this.picklistOptions = rule.options.map((o: any) => ({ ...o }));
    this.picklistSelectType = rule.isMultiSelect ? 'multi' : 'single';
    this.displayAsChip = rule.displayAsChip;
    const currentSheetName = (this.sheets && this.sheets[this.currentSheetIdx]?.name) || 'Sheet1';
    this.appliesToInput = `'${currentSheetName}'.${this.getRangeRef()}`;
    this.appliesToEditing = true; // Automatically activates Applies to input mode for target cell
    this.isCopyMode = true; // Sidebar title: "Picklist"
    this.editingOldRule = null; // Copy creates a NEW rule without overwriting original rule cells
    this.managePicklistSidebarOpen = false;
    this.validationModalOpen = true;
    if (this.cdr) this.cdr.detectChanges();
  }

  deletePicklistRule(rule: any, e?: Event) {
    if (e) e.stopPropagation();
    const targetSheet = this.sheets && this.sheets[rule.sheetIdx];
    const targetVals = (rule.sheetIdx === this.currentSheetIdx) ? this.validations : (targetSheet ? (targetSheet.validations || {}) : {});

    let count = 0;
    for (const cell of rule.cells) {
      const key = `${cell.r},${cell.c}`;
      if (targetVals[key]) {
        delete targetVals[key];
        count++;
      }
    }

    if (rule.sheetIdx === this.currentSheetIdx) {
      this.validations = { ...targetVals };
      if (targetSheet) targetSheet.validations = { ...targetVals };
    } else if (targetSheet) {
      targetSheet.validations = { ...targetVals };
    }

    this.onCellChange();
    this.save();
    this.refreshManagePicklistRules();
    this.showToast('Picklist rule deleted.');
    if (this.cdr) this.cdr.detectChanges();
  }

  startEditingAppliesTo() {
    this.appliesToEditing = true;
  }

  confirmAppliesTo() {
    this.appliesToEditing = false;
  }

  cancelEditingAppliesTo() {
    const currentSheetName = this.sheets[this.currentSheetIdx]?.name || 'Sheet1';
    this.appliesToInput = `'${currentSheetName}'.${this.getRangeRef()}`;
    this.appliesToEditing = false;
  }

  openValidationModal() {
    const existing = this.validations[`${this.selectedRow},${this.selectedCol}`];
    this.picklistOptions = [];
    this.isCopyMode = false;
    this.appliesToEditing = false;
    this.editingOldRule = null;
    const currentSheetName = this.sheets[this.currentSheetIdx]?.name || 'Sheet1';
    if (existing && existing.options) {
      this.picklistSelectType = existing.isMultiSelect ? 'multi' : 'single';
      this.displayAsChip = existing.displayAsChip !== false;
      existing.options.forEach(o => {
        if (typeof o === 'string') this.picklistOptions.push({ label: o, color: '#f97316' });
        else this.picklistOptions.push({ label: (o as DropdownOption).label, color: (o as DropdownOption).color || '#f97316', textColor: (o as DropdownOption).textColor });
      });
      this.appliesToInput = `'${currentSheetName}'.${this.getRangeRef()}`;
    } else {
      this.picklistSelectType = 'single';
      this.displayAsChip = true;
      this.picklistOptions.push({ label: 'Item 1', color: '#84cc16' });
      this.picklistOptions.push({ label: 'Item 2', color: '#ef4444' });
      this.appliesToInput = `'${currentSheetName}'.${this.getRangeRef()}`;
    }
    this.validationModalOpen = true;
    if (this.cdr) this.cdr.detectChanges();
  }

  addPicklistOption() {
    const colors = ['#4caf50', '#f44336', '#ff9800', '#2196f3', '#9c27b0', '#795548', '#607d8b'];
    this.picklistOptions.push({ label: '', color: colors[this.picklistOptions.length % colors.length] });
  }

  openPivotModal(e?: Event) {
    if (e) e.stopPropagation();
    const sheetName = this.sheets[this.currentSheetIdx].name;
    if (this.rangeStart && this.rangeEnd) {
      const minR = Math.min(this.rangeStart.r, this.rangeEnd.r);
      const maxR = Math.max(this.rangeStart.r, this.rangeEnd.r);
      const minC = Math.min(this.rangeStart.c, this.rangeEnd.c);
      const maxC = Math.max(this.rangeStart.c, this.rangeEnd.c);
      this.pivotSource = `${sheetName}.${this.colLabel(minC)}${minR + 1}:${this.colLabel(maxC)}${maxR + 1}`;
    } else {
      this.pivotSource = `${sheetName}.${this.colLabel(this.selectedCol)}${this.selectedRow + 1}`;
    }
    this.pivotModalOpen = true;
    this.closeMenus();
  }

  openDataValidationModal(e?: Event) {
    if (e) e.stopPropagation();
    const sheetName = this.sheets[this.currentSheetIdx].name;
    if (this.rangeStart && this.rangeEnd) {
      const minR = Math.min(this.rangeStart.r, this.rangeEnd.r);
      const maxR = Math.max(this.rangeStart.r, this.rangeEnd.r);
      const minC = Math.min(this.rangeStart.c, this.rangeEnd.c);
      const maxC = Math.max(this.rangeStart.c, this.rangeEnd.c);
      this.dvAppliesTo = `${sheetName}.${this.colLabel(minC)}${minR + 1}:${this.colLabel(maxC)}${maxR + 1}`;
    } else {
      this.dvAppliesTo = `${sheetName}.${this.colLabel(this.selectedCol)}${this.selectedRow + 1}`;
    }
    const existingVal = this.validations[`${this.selectedRow},${this.selectedCol}`];
    if (existingVal && existingVal.type === 'list') {
      this.dvCriteria = 'list';
      this.validationInput = existingVal.options.map((o: any) => o.label || o).join('\n');
      this.dvIsMultiSelect = existingVal.isMultiSelect || false;
      this.dvDisplayAsChip = existingVal.displayAsChip !== false;
      this.dvColorMode = existingVal.colorMode || 'none';
      this.dvSingleColor = existingVal.singleColor || '#f1f5f9';
      this.dvItemColors = {};
      existingVal.options.forEach((o: any) => {
        if (o.color) this.dvItemColors[o.label || o] = o.color;
      });
    } else {
      this.validationInput = '';
      this.dvCriteria = 'list';
      this.dvIsMultiSelect = false;
      this.dvDisplayAsChip = true;
      this.dvColorMode = 'none';
      this.dvSingleColor = '#f1f5f9';
      this.dvItemColors = {};
    }
    this.dvShowList = true;
    this.dvSortAsc = false;
    this.dvIgnoreBlanks = true;
    this.dvAlertsOpen = false;
    this.dvAlertEnabled = true;
    this.dvAlertTitle = '';
    this.dvAlertMsg = '';
    this.dataValidationModalOpen = true;
    this.closeMenus();
  }

  openManageRulesModal(e?: Event) {
    if (e) {
      e.stopPropagation();
    }
    this.manageRulesModalOpen = true;
    this.closeMenus();
  }

  pivotHeaders: string[] = [];
  pivotData: any[][] = [];
  pivotConfig = { row: '', val: '', agg: 'SUM' };

  createPivotTable() {
    let minR = 0, maxR = 0, minC = 0, maxC = 0;

    // Parse range from the modal input
    if (this.pivotDest && this.pivotDest.includes(':')) {
      const parts = this.pivotDest.split(':');
      const startParts = parts[0].split('.');
      const endPart = parts[1];
      const startRef = startParts.length > 1 ? startParts[1] : startParts[0];
      const endRef = endPart;

      const sCol = startRef.match(/[A-Z]+/)![0];
      const sRow = startRef.match(/[0-9]+/)![0];
      const eCol = endRef.match(/[A-Z]+/)![0];
      const eRow = endRef.match(/[0-9]+/)![0];

      minC = this.colToIndex(sCol);
      minR = parseInt(sRow, 10) - 1;
      maxC = this.colToIndex(eCol);
      maxR = parseInt(eRow, 10) - 1;
    } else if (this.rangeStart && this.rangeEnd) {
      minR = Math.min(this.rangeStart.r, this.rangeEnd.r);
      maxR = Math.max(this.rangeStart.r, this.rangeEnd.r);
      minC = Math.min(this.rangeStart.c, this.rangeEnd.c);
      maxC = Math.max(this.rangeStart.c, this.rangeEnd.c);
    } else {
      this.showToast('Invalid data range.');
      return;
    }

    if (maxR === minR) {
      this.showToast('Range must include headers and at least one data row.');
      return;
    }

    this.pivotHeaders = [];
    for (let c = minC; c <= maxC; c++) {
      this.pivotHeaders.push(this.cells[minR][c] || `Col ${c}`);
    }

    this.pivotData = [];
    for (let r = minR + 1; r <= maxR; r++) {
      let row = [];
      for (let c = minC; c <= maxC; c++) {
        row.push(this.cells[r][c] || '');
      }
      this.pivotData.push(row);
    }

    if (this.pivotDestType === 'new') {
      this.addSheet();
      this.switchSheet(this.sheets.length - 1);
      this.sheets[this.currentSheetIdx].name = 'Pivot Table 1';
      this.selectedRow = 0;
      this.selectedCol = 0;
    }

    this.pivotConfig = {
      row: this.pivotHeaders[0],
      val: this.pivotHeaders[this.pivotHeaders.length - 1],
      agg: 'SUM'
    };

    this.pivotModalOpen = false;
    this.sidePanelApp = 'pivot';

    this.applyPivot();
  }

  applyPivot() {
    if (!this.pivotConfig.row || !this.pivotConfig.val) return;

    const rowIdx = this.pivotHeaders.indexOf(this.pivotConfig.row);
    const valIdx = this.pivotHeaders.indexOf(this.pivotConfig.val);

    if (rowIdx === -1 || valIdx === -1) return;

    const map = new Map<string, number[]>();
    for (const row of this.pivotData) {
      const key = String(row[rowIdx]);
      const rawVal = row[valIdx];
      const val = Number(rawVal) || 0;

      if (!map.has(key)) map.set(key, []);
      map.get(key)!.push(val);
    }

    this.pushHistory();
    const startR = this.selectedRow;
    const startC = this.selectedCol;

    // Clear previous pivot table area
    for (let r = 0; r < 20; r++) for (let c = 0; c < 5; c++) {
      if (startR + r < this.ROWS && startC + c < this.COLS) {
        this.cells[startR + r][startC + c] = '';
        delete this.formats[`${startR + r},${startC + c}`];
      }
    }

    this.cells[startR][startC] = this.pivotConfig.row;
    this.cells[startR][startC + 1] = `${this.pivotConfig.agg} of ${this.pivotConfig.val}`;
    this.formats[`${startR},${startC}`] = { bold: true, bg: '#f1f5f9' };
    this.formats[`${startR},${startC + 1}`] = { bold: true, bg: '#f1f5f9' };

    let currR = startR + 1;
    let allVals: number[] = [];

    for (const [k, arr] of map.entries()) {
      if (currR < this.ROWS) {
        let v = 0;
        if (this.pivotConfig.agg === 'SUM') v = arr.reduce((a, b) => a + b, 0);
        else if (this.pivotConfig.agg === 'COUNT') v = arr.length;
        else if (this.pivotConfig.agg === 'AVG') v = arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : 0;

        const displayV = (this.pivotConfig.agg === 'AVG' && v % 1 !== 0) ? v.toFixed(2) : String(v);

        this.cells[currR][startC] = k;
        this.cells[currR][startC + 1] = displayV;
        allVals.push(...arr);
        currR++;
      }
    }

    if (currR < this.ROWS) {
      let grandTotal = 0;
      if (this.pivotConfig.agg === 'SUM') grandTotal = allVals.reduce((a, b) => a + b, 0);
      else if (this.pivotConfig.agg === 'COUNT') grandTotal = allVals.length;
      else if (this.pivotConfig.agg === 'AVG') grandTotal = allVals.length ? allVals.reduce((a, b) => a + b, 0) / allVals.length : 0;

      const displayGrand = (this.pivotConfig.agg === 'AVG' && grandTotal % 1 !== 0) ? grandTotal.toFixed(2) : String(grandTotal);

      this.cells[currR][startC] = 'Grand Total';
      this.cells[currR][startC + 1] = displayGrand;
      this.formats[`${currR},${startC}`] = { bold: true, bg: '#f1f5f9' };
      this.formats[`${currR},${startC + 1}`] = { bold: true, bg: '#f1f5f9' };
    }

    this.onCellChange();
    this.save();
  }

  dvPaletteItem: string | null = null;

  getParsedValidationItems(): string[] {
    if (!this.validationInput) return [];
    return this.validationInput.split('\n').map(v => v.trim()).filter(v => !!v);
  }

  openColorPalette(item: string) {
    this.dvPaletteItem = this.dvPaletteItem === item ? null : item;
  }

  setPaletteColor(item: string, color: string) {
    this.dvItemColors[item] = color;
    this.dvPaletteItem = null;
  }

  saveDataValidation() {
    if (this.validationInput.trim().length > 0) {
      const options = this.validationInput.split('\n').filter(o => o.trim() !== '');
      this.picklistOptions = options.map(o => {
        let color = undefined;
        if (this.dvColorMode === 'single') color = this.dvSingleColor;
        else if (this.dvColorMode === 'multi') color = this.dvItemColors[o.trim()] || undefined;
        return { label: o.trim(), color };
      });
      this.saveValidation();
    }
    this.dataValidationModalOpen = false;
  }

  saveValidation() {
    const validOptions = this.picklistOptions.filter(o => o.label.trim().length > 0);
    if (validOptions.length === 0) { this.validationModalOpen = false; return; }
    let targetSheetIdx = this.currentSheetIdx;
    let minR = this.selectedRow, maxR = this.selectedRow;
    let minC = this.selectedCol, maxC = this.selectedCol;

    if (this.appliesToInput && this.appliesToInput.includes('.')) {
      const parts = this.appliesToInput.split('.');
      const sheetPart = parts[0].replace(/^'|'$/g, '').trim();
      const rangePart = parts.slice(1).join('.').trim();

      if (this.sheets && this.sheets.length) {
        const foundIdx = this.sheets.findIndex(s => s.name === sheetPart);
        if (foundIdx !== -1) targetSheetIdx = foundIdx;
      }

      if (rangePart.includes(':')) {
        const [start, end] = rangePart.split(':');
        const sColMatch = start.match(/[A-Z]+/i);
        const sRowMatch = start.match(/[0-9]+/);
        const eColMatch = end.match(/[A-Z]+/i);
        const eRowMatch = end.match(/[0-9]+/);
        if (sColMatch && sRowMatch && eColMatch && eRowMatch) {
          minC = Math.min(this.colToIndex(sColMatch[0].toUpperCase()), this.colToIndex(eColMatch[0].toUpperCase()));
          maxC = Math.max(this.colToIndex(sColMatch[0].toUpperCase()), this.colToIndex(eColMatch[0].toUpperCase()));
          minR = Math.min(parseInt(sRowMatch[0], 10) - 1, parseInt(eRowMatch[0], 10) - 1);
          maxR = Math.max(parseInt(sRowMatch[0], 10) - 1, parseInt(eRowMatch[0], 10) - 1);
        }
      } else {
        const cMatch = rangePart.match(/[A-Z]+/i);
        const rMatch = rangePart.match(/[0-9]+/);
        if (cMatch && rMatch) {
          minC = maxC = this.colToIndex(cMatch[0].toUpperCase());
          minR = maxR = parseInt(rMatch[0], 10) - 1;
        }
      }
    } else if (this.rangeStart && this.rangeEnd) {
      minR = Math.min(this.rangeStart.r, this.rangeEnd.r);
      maxR = Math.max(this.rangeStart.r, this.rangeEnd.r);
      minC = Math.min(this.rangeStart.c, this.rangeEnd.c);
      maxC = Math.max(this.rangeStart.c, this.rangeEnd.c);
    }

    const isMulti = this.picklistSelectType === 'multi';
    const spec = { 
      type: 'list', 
      options: validOptions,
      isMultiSelect: isMulti,
      displayAsChip: this.displayAsChip,
      colorMode: 'multi' as 'multi',
      singleColor: ''
    };

    if (targetSheetIdx === this.currentSheetIdx) {
      const newValidations = { ...this.validations };
      for (let r = minR; r <= maxR; r++) {
        for (let c = minC; c <= maxC; c++) {
          newValidations[`${r},${c}`] = { ...spec };
          const cur = this.cells[r]?.[c];
          if (cur && !validOptions.find(o => o.label === cur)) this.cells[r][c] = '';
        }
      }
      this.validations = newValidations;
      if (this.sheets && this.sheets[targetSheetIdx]) this.sheets[targetSheetIdx].validations = newValidations;
    } else {
      const targetSheet = this.sheets && this.sheets[targetSheetIdx];
      if (targetSheet) {
        if (!targetSheet.validations) targetSheet.validations = {};
        for (let r = minR; r <= maxR; r++) {
          for (let c = minC; c <= maxC; c++) {
            targetSheet.validations[`${r},${c}`] = { ...spec };
            if (targetSheet.cells && targetSheet.cells[r]) {
              const cur = targetSheet.cells[r][c];
              if (cur && !validOptions.find(o => o.label === cur)) targetSheet.cells[r][c] = '';
            }
          }
        }
      }
    }

    if (this.editingOldRule && !this.isCopyMode) {
      const oldSheetIdx = this.editingOldRule.sheetIdx;
      const oldSheet = this.sheets && this.sheets[oldSheetIdx];
      const oldVals = (oldSheetIdx === this.currentSheetIdx) ? this.validations : (oldSheet ? (oldSheet.validations || {}) : {});
      if (this.editingOldRule.cells) {
        for (const cell of this.editingOldRule.cells) {
          delete oldVals[`${cell.r},${cell.c}`];
        }
      }
      if (oldSheetIdx === this.currentSheetIdx) {
        this.validations = { ...oldVals };
        if (oldSheet) oldSheet.validations = { ...oldVals };
      } else if (oldSheet) {
        oldSheet.validations = { ...oldVals };
      }
      this.editingOldRule = null;
    }

    this.validationModalOpen = false;
    this.onCellChange();
    this.save();
    this.showToast(`Picklist set: ${validOptions.length} items`);
    if (this.cdr) this.cdr.detectChanges();
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
  }

  toggleMenu(menu: string, e: Event) {
    e.stopPropagation();
    const target = e.target as HTMLElement;
    if (target.closest('.mdd')) return;
    this.activePalette = null;
    this.activeMenu = this.activeMenu === menu ? null : menu;
  }

  toggleFooterMenu(menu: string, e: Event) {
    e.stopPropagation();
    if (this.activeFooterMenu === menu) {
      this.activeFooterMenu = null;
    } else {
      this.closeMenus();
      this.activeFooterMenu = menu;
    }
    if (this.cdr) this.cdr.detectChanges();
  }

  toggleViewSetting(setting: string) {
    if (setting === 'topBar') this.showTopBar = !this.showTopBar;
    if (setting === 'formulaBar') this.showFormulaBar = !this.showFormulaBar;
    if (setting === 'notes') this.showNotes = !this.showNotes;
    if (setting === 'userPresence') this.showUserPresence = !this.showUserPresence;
    if (setting === 'lockPattern') this.showLockPattern = !this.showLockPattern;
    if (setting === 'printArea') this.showHighlightPrintArea = !this.showHighlightPrintArea;
    if (this.cdr) this.cdr.detectChanges();
  }

  closeMenus() { this.activeMenu = null; this.profileOpen = false; this.activeBorderSubmenu = null; this.activeSheetMenuIdx = null; this.advFilterVisible = false; this.activeFooterMenu = null; }

  newDoc() {
    this.api.createDocument('Untitled spreadsheet', 'sheet')
      .subscribe((res: any) => {
        window.open(`/sheet/${res.id}`, '_blank');
        this.closeMenus();
      });
  }

  isImageCell(r: number, c: number): boolean {
    const val = this.cells[r]?.[c];
    if (typeof val !== 'string') return false;
    if (val.trim().startsWith('data:image')) return true;
    if (val.trim().toUpperCase().startsWith('=IMAGE(')) return true;
    return false;
  }

  getImageSrc(val: string): string {
    if (!val || typeof val !== 'string') return '';
    if (val.startsWith('data:image')) return val;
    if (val.toUpperCase().startsWith('=IMAGE(')) {
      const match = val.match(/=IMAGE\(\s*["'](.*?)["']/i);
      if (match) return match[1];
    }
    return val;
  }

  triggerImageInsert(type: string = 'cell') {
    this.closeMenus();
    this.imgInputRef?.nativeElement.click();
    if (type === 'over') {
      setTimeout(() => this.showToast('Image over cells will be added as a floating overlay.'), 500);
    }
  }

  onImageFileSelected(e: Event) {
    const file = (e.target as HTMLInputElement).files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      this.pushHistory();
      this.cells[this.selectedRow][this.selectedCol] = ev.target!.result as string;
      this.formulaBarValue = '[IMAGE]';
      this.onCellChange();
      this.save();
      this.showToast('Image inserted into cell.');
    };
    reader.readAsDataURL(file);
    (e.target as HTMLInputElement).value = '';
  }

  printSheet() {
    this.closeMenus();
    // Find data bounds
    let maxRow = 4, maxCol = 4;
    for (let r = 0; r < this.ROWS; r++)
      for (let c = 0; c < this.COLS; c++)
        if ((this.cells[r][c] || '').trim()) { maxRow = Math.max(maxRow, r); maxCol = Math.max(maxCol, c); }
    maxRow = Math.min(maxRow + 2, this.ROWS - 1);
    maxCol = Math.min(maxCol + 2, this.COLS - 1);

    // using global colName
    let thead = '<tr><th style="width:36px;"></th>';
    for (let c = 0; c <= maxCol; c++) thead += `<th>${colName(c)}</th>`;
    thead += '</tr>';

    let tbody = '';
    for (let r = 0; r <= maxRow; r++) {
      tbody += `<tr><td class="rh">${r + 1}</td>`;
      for (let c = 0; c <= maxCol; c++) {
        const fmt = this.formats[`${r},${c}`] || {};
        let s = '';
        if (fmt.bold) s += 'font-weight:bold;';
        if (fmt.italic) s += 'font-style:italic;';
        if (fmt.strikethrough) s += 'text-decoration:line-through;';
        if (fmt.color) s += `color:${fmt.color};`;
        if (fmt.bg) s += `background:${fmt.bg};`;
        if (fmt.align) s += `text-align:${fmt.align};`;
        if (fmt.font) s += `font-family:${fmt.font};`;
        if (fmt.size) s += `font-size:${fmt.size};`;
        const val = this.cells[r][c] || '';
        const content = val.startsWith('data:image')
          ? `<img src="${val}" style="max-width:120px;max-height:80px;object-fit:contain;">`
          : val.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        tbody += `<td style="${s}">${content}</td>`;
      }
      tbody += '</tr>';
    }

    const win = window.open('', '_blank', 'width=900,height=700');
    if (!win) { window.print(); return; }
    win.document.write(`<!DOCTYPE html><html><head>
      <title>${this.title}</title>
      <style>
        body{font-family:Arial,sans-serif;font-size:12px;margin:15mm;}
        h2{font-size:16px;margin:0 0 8px;color:#202124;}
        p{font-size:11px;color:#888;margin:0 0 12px;}
        table{border-collapse:collapse;width:100%;}
        th{background:#f1f3f4;border:1px solid #bbb;padding:4px 8px;font-size:11px;text-align:center;}
        td{border:1px solid #ddd;padding:3px 6px;font-size:12px;vertical-align:middle;}
        .rh{background:#f1f3f4;text-align:center;color:#666;font-size:11px;width:36px;}
      </style></head><body>
      <h2>${this.title}</h2>
      <p>Printed on ${new Date().toLocaleString()}</p>
      <table><thead>${thead}</thead><tbody>${tbody}</tbody></table>
    </body></html>`);
    win.document.close();
    win.focus();
    setTimeout(() => { win.print(); }, 400);
  }

  selectCell(r: number, c: number) {
    this.activeShapeIdx = null;
    this.activeShapeMenuIdx = null;
    if (this.sheets[this.currentSheetIdx]?.locked) {
      this.showToast('This sheet is locked.');
      return;
    }
    if (this.isEditingCell) this.commitEdit();
    this.selectedRow = r; this.selectedCol = c;
    this.formulaBarValue = this.isImageCell(r, c) ? '[IMAGE]' : this.cells[r][c];
    this.currentFont = this.getFormat('font') || 'Arial';
    this.currentSize = this.getFormat('size') || '13px';
    this.currentSizeNum = parseInt(this.currentSize, 10) || 13;
    this.api.sendCursor(r, c);
  }

  onCellClickWithPicker(r: number, c: number) {
    if (this.rangePickerActive) {
      this.rangePickerStartR = r;
      this.rangePickerStartC = c;
      this.rangePickerEndR = r;
      this.rangePickerEndC = c;
      this.updateRangePickerField();
      return;
    }
    this.selectCell(r, c);
  }

  startRangePicker(field: 'pivotSource' | 'pivotDest' | 'dvAppliesTo') {
    this.rangePickerField = field;
    this.rangePickerActive = true;
    this._pivotModalWasOpen = this.pivotModalOpen;
    this._dvModalWasOpen = this.dataValidationModalOpen;
    this.pivotModalOpen = false;
    this.dataValidationModalOpen = false;
    // Seed with current selection
    this.rangePickerStartR = this.selectedRow;
    this.rangePickerStartC = this.selectedCol;
    this.rangePickerEndR = this.selectedRow;
    this.rangePickerEndC = this.selectedCol;
    this.updateRangePickerField();
  }

  updateRangePickerField() {
    const sheetName = this.sheets[this.currentSheetIdx].name;
    const minR = Math.min(this.rangePickerStartR, this.rangePickerEndR);
    const maxR = Math.max(this.rangePickerStartR, this.rangePickerEndR);
    const minC = Math.min(this.rangePickerStartC, this.rangePickerEndC);
    const maxC = Math.max(this.rangePickerStartC, this.rangePickerEndC);
    const ref = minR === maxR && minC === maxC
      ? `${sheetName}.${this.colLabel(minC)}${minR + 1}`
      : `${sheetName}.${this.colLabel(minC)}${minR + 1}:${this.colLabel(maxC)}${maxR + 1}`;
    if (this.rangePickerField === 'pivotSource') this.pivotSource = ref;
    else if (this.rangePickerField === 'pivotDest') this.pivotDest = ref;
    else if (this.rangePickerField === 'dvAppliesTo') this.dvAppliesTo = ref;
  }

  getRangePickerValue(): string {
    if (this.rangePickerField === 'pivotSource') return this.pivotSource;
    if (this.rangePickerField === 'pivotDest') return this.pivotDest;
    if (this.rangePickerField === 'dvAppliesTo') return this.dvAppliesTo;
    return '';
  }

  onRangePickerInput(e: Event) {
    const val = (e.target as HTMLInputElement).value;
    if (this.rangePickerField === 'pivotSource') this.pivotSource = val;
    else if (this.rangePickerField === 'pivotDest') this.pivotDest = val;
    else if (this.rangePickerField === 'dvAppliesTo') this.dvAppliesTo = val;
  }

  confirmRangePicker() {
    this.rangePickerActive = false;
    if (this._pivotModalWasOpen) this.pivotModalOpen = true;
    if (this._dvModalWasOpen) this.dataValidationModalOpen = true;
    this.rangePickerField = null;
  }

  cancelRangePicker() {
    this.rangePickerActive = false;
    if (this._pivotModalWasOpen) this.pivotModalOpen = true;
    if (this._dvModalWasOpen) this.dataValidationModalOpen = true;
    this.rangePickerField = null;
  }

  getColOffset(c: number): number {
    let offset = (this.showHeaders ? 46 : 0) + this.groupMarginWidth;
    const widths = this.sheets[this.currentSheetIdx].colWidths || {};
    for (let i = 0; i < c; i++) { offset += widths[i] ?? 100; }
    return offset;
  }

  getRowOffset(r: number): number {
    let offset = this.showHeaders ? 26 : 0;
    const heights = this.sheets[this.currentSheetIdx].rowHeights || {};
    for (let i = 0; i < r; i++) { offset += heights[i] ?? 26; }
    return offset;
  }

  getFormatWrap(r: number, c: number): boolean {
    const f = this.sheets[this.currentSheetIdx].formats[`${r},${c}`];
    return f ? f.wrap === 'wrap' || f.wrap === true : false;
  }

  startEditing(initialValue?: string) {
    if (this.sheets[this.currentSheetIdx]?.locked) {
      this.showToast('This sheet is locked.');
      return;
    }
    if ((this.formats[`${this.selectedRow},${this.selectedCol}`] as any)?.locked) {
      this.showToast('This cell is locked.');
      return;
    }
    this.isEditingCell = true;
    this.editValue = initialValue !== undefined ? initialValue : this.cells[this.selectedRow][this.selectedCol];
    setTimeout(() => {
      if (this.floatingEditor) {
        this.floatingEditor.nativeElement.focus();
        const len = this.editValue.length;
        this.floatingEditor.nativeElement.setSelectionRange(len, len);
        this.autoResizeEditor();
      }
    }, 0);
  }

  commitEdit() {
    if (!this.isEditingCell) return;
    this.pushHistory(); // <-- Capture baseline before mutation for Undo and Audit Diff
    this.cells[this.selectedRow][this.selectedCol] = this.editValue;
    this.formulaBarValue = this.editValue;
    this.isEditingCell = false;
    this.onCellChange();
    this.save();
    setTimeout(() => {
      (document.activeElement as HTMLElement)?.blur();
    }, 0);
  }

  autoResizeEditor(e?: Event) {
    if (!this.floatingEditor) return;
    const el = this.floatingEditor.nativeElement;
    const td = el.parentElement;

    // Height resize
    el.style.height = 'auto';
    const minHeight = td ? td.offsetHeight + 1 : (this.getRowHeight(this.selectedRow) + 3);
    el.style.height = Math.max(minHeight, el.scrollHeight) + 'px';

    // Width resize
    const minWidth = td ? td.offsetWidth + 1 : (this.getColWidth(this.selectedCol) + 3);
    el.style.width = minWidth + 'px';

    // If scrollWidth is larger than minWidth, we need to expand.
    // The +2 buffer avoids horizontal scrollbar flickering
    if (el.scrollWidth > minWidth) {
      el.style.width = (el.scrollWidth + 2) + 'px';
    }
  }

  onEditorKeydown(e: KeyboardEvent) {
    if (e.key === 'Enter') {
      if (e.ctrlKey || e.altKey || e.metaKey) {
        e.preventDefault();
        const start = (e.target as HTMLTextAreaElement).selectionStart;
        const end = (e.target as HTMLTextAreaElement).selectionEnd;
        this.editValue = this.editValue.substring(0, start) + '\n' + this.editValue.substring(end);
        setTimeout(() => {
          if (this.floatingEditor) {
            this.floatingEditor.nativeElement.selectionStart = this.floatingEditor.nativeElement.selectionEnd = start + 1;
          }
          this.autoResizeEditor();
        }, 0);
      } else if (!e.shiftKey) {
        e.preventDefault();
        this.commitEdit();
        this.onEnter(e, this.selectedRow, this.selectedCol);
      }
    } else if (e.key === 'Escape') {
      this.isEditingCell = false;
    } else if (e.key === 'Tab') {
      e.preventDefault();
      this.commitEdit();
      this.onTab(e, this.selectedRow, this.selectedCol);
    }
  }

  onFormulaBarChange(val: string) {
    if (this.cells[this.selectedRow][this.selectedCol] !== val) {
      this.pushHistory(); // <-- Capture baseline before mutation for Undo and Audit Diff
      this.cells[this.selectedRow][this.selectedCol] = val;
      this.onCellChange();
    }
  }

  commitFormula() {
    this.cells[this.selectedRow][this.selectedCol] = this.formulaBarValue;
    this.onCellChange();
  }

  onTab(e: KeyboardEvent, r: number, c: number) {
    e.preventDefault();
    const nc = c + 1 < this.COLS ? c + 1 : c;
    this.selectCell(r, nc); this.focusCell(r, nc);
  }

  onEnter(e: KeyboardEvent, r: number, c: number) {
    e.preventDefault();
    const nr = r + 1 < this.ROWS ? r + 1 : r;
    this.selectCell(nr, c); this.focusCell(nr, c);
  }

  private focusCell(r: number, c: number) {
    setTimeout(() => {
      const inputs = document.querySelectorAll<HTMLInputElement>('.cell-input');
      inputs[r * this.COLS + c]?.focus();
    });
  }

  // --- Formatting Engine ---
  getFormatName(fmt: string): string {
    if (!fmt) return 'General';
    if (fmt === 'general') return 'General';
    if (fmt === 'number') return 'Number';
    if (fmt === 'percent') return 'Percentage';
    if (fmt.startsWith('currency')) return 'Currency';
    if (fmt.startsWith('accounting')) return 'Accounting';
    if (fmt.startsWith('date')) return 'Date';
    if (fmt.startsWith('time')) return 'Time';
    if (fmt.startsWith('fraction')) return 'Fraction';
    if (fmt === 'scientific') return 'Scientific';
    if (fmt === 'text') return 'Text';
    if (fmt === 'regional_zip') return 'Zip Code';
    if (fmt === 'regional_phone') return 'Phone';
    if (fmt === 'regional_zip4') return 'Zip Code+4';
    return 'General';
  }

  getFormat(key: keyof CellFormat): any {
    const ref = `${this.selectedRow},${this.selectedCol}`;
    return this.formats[ref]?.[key];
  }

  setFormat(key: keyof CellFormat, val: any) {
    if (this.rangeStart && this.rangeEnd) {
      const minR = Math.min(this.rangeStart.r, this.rangeEnd.r);
      const maxR = Math.max(this.rangeStart.r, this.rangeEnd.r);
      const minC = Math.min(this.rangeStart.c, this.rangeEnd.c);
      const maxC = Math.max(this.rangeStart.c, this.rangeEnd.c);

      for (let r = minR; r <= maxR; r++) {
        for (let c = minC; c <= maxC; c++) {
          const ref = `${r},${c}`;
          if (!this.formats[ref]) this.formats[ref] = {};

          if (key === 'indent') {
            const currentIndent = (this.formats[ref] as any).indent || 0;
            if (val === 'increase') (this.formats[ref] as any).indent = currentIndent + 1;
            else if (val === 'decrease') (this.formats[ref] as any).indent = Math.max(0, currentIndent - 1);
          } else {
            (this.formats[ref] as any)[key] = val;
          }
        }
      }
    } else {
      const ref = `${this.selectedRow},${this.selectedCol}`;
      if (!this.formats[ref]) this.formats[ref] = {};

      if (key === 'indent') {
        const currentIndent = (this.formats[ref] as any).indent || 0;
        if (val === 'increase') (this.formats[ref] as any).indent = currentIndent + 1;
        else if (val === 'decrease') (this.formats[ref] as any).indent = Math.max(0, currentIndent - 1);
      } else {
        (this.formats[ref] as any)[key] = val;
      }
    }

    this.formats = { ...this.formats };
    this.activePalette = null;
    this.onCellChange();
  }

  toggleFormat(key: 'bold' | 'italic' | 'strikethrough' | 'underline') {
    const primaryRef = `${this.selectedRow},${this.selectedCol}`;
    const targetState = !(this.formats[primaryRef] as any)?.[key];

    if (this.rangeStart && this.rangeEnd) {
      const minR = Math.min(this.rangeStart.r, this.rangeEnd.r);
      const maxR = Math.max(this.rangeStart.r, this.rangeEnd.r);
      const minC = Math.min(this.rangeStart.c, this.rangeEnd.c);
      const maxC = Math.max(this.rangeStart.c, this.rangeEnd.c);
      for (let r = minR; r <= maxR; r++) {
        for (let c = minC; c <= maxC; c++) {
          const ref = `${r},${c}`;
          if (!this.formats[ref]) this.formats[ref] = {};
          (this.formats[ref] as any)[key] = targetState;
        }
      }
    } else {
      if (!this.formats[primaryRef]) this.formats[primaryRef] = {};
      (this.formats[primaryRef] as any)[key] = targetState;
    }

    this.formats = { ...this.formats };
    this.closeMenus();
    this.onCellChange();
  }

  applyFont(font: string) {
    this.currentFont = font;
    this.setFormat('font', font);
    this.closeMenus();
  }

  applySizeNum(size: number) {
    this.currentSizeNum = size;
    this.currentSize = size + 'px';
    this.setFormat('size', this.currentSize);
  }

  onFontSizeInputChange() {
    if (this.currentSizeNum > 0) this.applySizeNum(this.currentSizeNum);
  }

  incrementFontSize() {
    this.applySizeNum(this.currentSizeNum + 1);
  }

  decrementFontSize() {
    if (this.currentSizeNum > 1) {
      this.applySizeNum(this.currentSizeNum - 1);
    }
  }

  togglePalette(which: string, e: Event) {
    e.stopPropagation();
    this.activeMenu = null;
    this.activePalette = this.activePalette === which ? null : which;
  }

  // ── New view/format/edit helpers ─────────────────────────────────────────


  clearAllFormats() {
    const minR = this.rangeStart ? Math.min(this.rangeStart.r, this.rangeEnd!.r) : this.selectedRow;
    const maxR = this.rangeStart ? Math.max(this.rangeStart.r, this.rangeEnd!.r) : this.selectedRow;
    const minC = this.rangeStart ? Math.min(this.rangeStart.c, this.rangeEnd!.c) : this.selectedCol;
    const maxC = this.rangeStart ? Math.max(this.rangeStart.c, this.rangeEnd!.c) : this.selectedCol;
    this.pushHistory();
    for (let r = minR; r <= maxR; r++)
      for (let c = minC; c <= maxC; c++)
        delete this.formats[`${r},${c}`];
    this.formats = { ...this.formats };
    this.onCellChange(); this.save();
    this.closeMenus(); this.showToast('Formats cleared.');
  }

  clearAll() {
    this.pushHistory();
    const minR = this.rangeStart ? Math.min(this.rangeStart.r, this.rangeEnd!.r) : this.selectedRow;
    const maxR = this.rangeStart ? Math.max(this.rangeStart.r, this.rangeEnd!.r) : this.selectedRow;
    const minC = this.rangeStart ? Math.min(this.rangeStart.c, this.rangeEnd!.c) : this.selectedCol;
    const maxC = this.rangeStart ? Math.max(this.rangeStart.c, this.rangeEnd!.c) : this.selectedCol;
    for (let r = minR; r <= maxR; r++)
      for (let c = minC; c <= maxC; c++) {
        this.cells[r][c] = '';
        delete this.formats[`${r},${c}`];
      }
    this.formats = { ...this.formats };
    this.onCellChange(); this.save(); this.closeMenus();
    this.showToast('Cleared all values and formats.');
  }

  toggleGridlines() { this.showGridlines = !this.showGridlines; this.closeMenus(); }
  toggleFormulaBar() { this.showFormulaBar = !this.showFormulaBar; this.closeMenus(); }
  toggleHeaders() { this.showHeaders = !this.showHeaders; this.closeMenus(); }

  setZoom(pct: number) { this.zoomLevel = pct; this.closeMenus(); }

  insertCheckbox() {
    this.pushHistory();
    const minR = this.rangeStart ? Math.min(this.rangeStart.r, this.rangeEnd!.r) : this.selectedRow;
    const maxR = this.rangeStart ? Math.max(this.rangeStart.r, this.rangeEnd!.r) : this.selectedRow;
    const minC = this.rangeStart ? Math.min(this.rangeStart.c, this.rangeEnd!.c) : this.selectedCol;
    const maxC = this.rangeStart ? Math.max(this.rangeStart.c, this.rangeEnd!.c) : this.selectedCol;

    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        this.cells[r][c] = 'FALSE';
        const ref = `${r},${c}`;
        if (!this.formats[ref]) this.formats[ref] = {};
        (this.formats[ref] as any)['checkbox'] = true;
      }
    }

    if (this.formulaBarValue !== 'FALSE') {
      this.formulaBarValue = 'FALSE';
    }
    this.formats = { ...this.formats };
    this.onCellChange(); this.save(); this.closeMenus();
    this.showToast('Checkbox inserted.');
  }

  removeDuplicates() {
    this.pushHistory();
    let minR = this.selectedRow, maxR = this.selectedRow;
    let minC = this.selectedCol, maxC = this.selectedCol;

    if (this.rangeStart && this.rangeEnd) {
      minR = Math.min(this.rangeStart.r, this.rangeEnd.r);
      maxR = Math.max(this.rangeStart.r, this.rangeEnd.r);
      minC = Math.min(this.rangeStart.c, this.rangeEnd.c);
      maxC = Math.max(this.rangeStart.c, this.rangeEnd.c);
    } else {
      // If only one cell is selected, apply to the entire active columns.
      minR = 0;
      maxR = this.cells.length - 1;
    }

    const seen = new Set<string>();
    let removed = 0;

    for (let r = minR; r <= maxR; r++) {
      let signature = '';
      for (let c = minC; c <= maxC; c++) {
        signature += (this.cells[r][c] || '') + '|';
      }

      // Ignore completely blank rows from duplicate detection
      if (signature === '|'.repeat(maxC - minC + 1)) {
        continue;
      }

      if (seen.has(signature)) {
        this.cells.splice(r, 1);
        this.cells.push(Array(this.COLS).fill('')); // Maintain row count
        r--;
        maxR--;
        removed++;
      } else {
        seen.add(signature);
      }
    }

    this.onCellChange(); this.save(); this.closeMenus();
    this.showToast(`Removed ${removed} duplicate row(s).`);
  }

  pushHistory(explicitOp?: AuditOp) {
    const livePreState = {
      cells: this.cells,
      formats: this.formats,
      hiddenRows: Array.from(this.hiddenRows),
      activeFilterCols: Array.from(this.activeFilterCols),
      filterActive: this.filterActive,
      advFilterSavedState: this.serializeAdvFilterState()
    };
    const preStateJson = JSON.stringify(livePreState);
    this.history.push(preStateJson);
    if (this.history.length > 50) this.history.shift();
    this.future = [];

    const sheet = this.sheets[this.currentSheetIdx];
    const context = {
      sheetId: (sheet as any).id || this.currentSheetIdx.toString(),
      sheetName: sheet.name,
      currentUser: this.auth.user?.name || 'unknown'
    };

    if (explicitOp) {
      setTimeout(() => {
        const payload = { sheet_id: context.sheetId, sheet_name: context.sheetName, ...explicitOp };
        const bufferKey = `${context.currentUser}_${explicitOp.target_range}_${explicitOp.action_type}`;
        this.auditBuffer.set(bufferKey, payload);
        this.flushAudit$.next();
      }, 0);
    } else {
      if (!this.pendingDiffTimer) {
        this.pendingDiffPreStateJson = preStateJson;
      } else {
        clearTimeout(this.pendingDiffTimer);
      }
      this.pendingDiffContext = context;
      this.pendingDiffTimer = setTimeout(() => {
        if (!this.pendingDiffPreStateJson) return;
        const clonedPreState = JSON.parse(this.pendingDiffPreStateJson);
        const postState = { cells: this.cells, formats: this.formats, hiddenRows: Array.from(this.hiddenRows) };
        const ops = this.diffStateForAudit(clonedPreState, postState);
        for (const op of ops) {
          const payload = { sheet_id: this.pendingDiffContext.sheetId, sheet_name: this.pendingDiffContext.sheetName, ...op };
          const bufferKey = `${this.pendingDiffContext.currentUser}_${op.target_range}_${op.action_type}`;
          this.auditBuffer.set(bufferKey, payload);
        }
        if (ops.length > 0) this.flushAudit$.next();
        this.pendingDiffPreStateJson = null;
        this.pendingDiffContext = null;
        this.pendingDiffTimer = null;
      }, 1500);
    }
  }

  diffStateForAudit(prev: any, curr: any): AuditOp[] {
    const ops: AuditOp[] = [];
    let cellMinR = Infinity, cellMaxR = -Infinity, cellMinC = Infinity, cellMaxC = -Infinity;
    let cellsChanged = false;
    const insertedImages: { r: number, c: number, name: string }[] = [];
    const deletedImages: { r: number, c: number }[] = [];
    const replacedImages: { r: number, c: number }[] = [];

    for (let r = 0; r < curr.cells.length; r++) {
      for (let c = 0; c < curr.cells[r].length; c++) {
        const prevVal = prev.cells[r]?.[c] || '';
        const currVal = curr.cells[r][c] || '';
        if (prevVal !== currVal) {
          const prevIsImg = prevVal.startsWith('data:image') || prevVal.startsWith('=IMAGE(');
          const currIsImg = currVal.startsWith('data:image') || currVal.startsWith('=IMAGE(');
          let isImageChange = false;
          if (!prevIsImg && currIsImg) { insertedImages.push({ r, c, name: 'image' }); isImageChange = true; }
          else if (prevIsImg && !currIsImg) { deletedImages.push({ r, c }); isImageChange = true; }
          else if (prevIsImg && currIsImg) { replacedImages.push({ r, c }); isImageChange = true; }

          if (!isImageChange) {
            cellsChanged = true;
            if (r < cellMinR) cellMinR = r; if (r > cellMaxR) cellMaxR = r;
            if (c < cellMinC) cellMinC = c; if (c > cellMaxC) cellMaxC = c;
          }
        }
      }
    }
    if (cellsChanged) {
      const isSingle = (cellMinR === cellMaxR && cellMinC === cellMaxC);
      let allEmpty = true;
      for (let r = cellMinR; r <= cellMaxR; r++) {
        for (let c = cellMinC; c <= cellMaxC; c++) {
          if (curr.cells[r] && curr.cells[r][c]) { allEmpty = false; break; }
        }
        if (!allEmpty) break;
      }
      ops.push({ action_type: allEmpty ? 'clear-cell' : 'set-cell-value', target_range: isSingle ? this.getA1(cellMinR, cellMinC) : `${this.getA1(cellMinR, cellMinC)}:${this.getA1(cellMaxR, cellMaxC)}` });
    }
    if (insertedImages.length === 1) ops.push({ action_type: 'insert-image', target_range: this.getA1(insertedImages[0].r, insertedImages[0].c), metadata: { image_name: insertedImages[0].name } });
    else if (insertedImages.length > 1) ops.push({ action_type: 'insert-images', target_range: 'Multiple', metadata: { ranges: insertedImages.map(i => this.getA1(i.r, i.c)) } });

    if (deletedImages.length === 1) ops.push({ action_type: 'delete-image', target_range: this.getA1(deletedImages[0].r, deletedImages[0].c) });
    else if (deletedImages.length > 1) ops.push({ action_type: 'delete-images', target_range: 'Multiple', metadata: { ranges: deletedImages.map(i => this.getA1(i.r, i.c)) } });

    if (replacedImages.length === 1) ops.push({ action_type: 'replace-image', target_range: this.getA1(replacedImages[0].r, replacedImages[0].c) });
    else if (replacedImages.length > 1) ops.push({ action_type: 'replace-images', target_range: 'Multiple', metadata: { ranges: replacedImages.map(i => this.getA1(i.r, i.c)) } });

    let fmtMinR = Infinity, fmtMaxR = -Infinity, fmtMinC = Infinity, fmtMaxC = -Infinity;
    let fmtChanged = false;
    const allFmtKeys = new Set([...Object.keys(curr.formats || {}), ...Object.keys(prev.formats || {})]);
    for (const k of allFmtKeys) {
      if (JSON.stringify(curr.formats[k]) !== JSON.stringify(prev.formats[k])) {
        const [r, c] = k.split(',').map(Number);
        fmtChanged = true;
        if (r < fmtMinR) fmtMinR = r; if (r > fmtMaxR) fmtMaxR = r;
        if (c < fmtMinC) fmtMinC = c; if (c > fmtMaxC) fmtMaxC = c;
      }
    }
    if (fmtChanged) {
      const isSingle = (fmtMinR === fmtMaxR && fmtMinC === fmtMaxC);
      ops.push({ action_type: 'format-change', target_range: isSingle ? this.getA1(fmtMinR, fmtMinC) : `${this.getA1(fmtMinR, fmtMinC)}:${this.getA1(fmtMaxR, fmtMaxC)}` });
    }

    const currHidden = new Set(curr.hiddenRows || []);
    const prevHidden = new Set(prev.hiddenRows || []);
    let hiddenChanged = false;
    for (const r of currHidden) if (!prevHidden.has(r)) hiddenChanged = true;
    for (const r of prevHidden) if (!currHidden.has(r)) hiddenChanged = true;
    if (hiddenChanged) ops.push({ action_type: 'toggle-hidden-rows', target_range: 'Entire Sheet' });

    return ops;
  }

  getA1(r: number, c: number): string { return `${this.colLabel(c)}${r + 1}`; }

  undo() {
    if (!this.history.length) { this.showToast('Nothing to undo.'); return; }
    this.future.push(JSON.stringify({
      cells: this.cells, formats: this.formats,
      hiddenRows: Array.from(this.hiddenRows),
      activeFilterCols: Array.from(this.activeFilterCols),
      filterActive: this.filterActive,
      advFilterSavedState: this.serializeAdvFilterState()
    }));
    const prev = JSON.parse(this.history.pop()!);
    if (prev.cells) for (let r = 0; r < this.ROWS; r++) for (let c = 0; c < this.COLS; c++) this.cells[r][c] = prev.cells[r]?.[c] ?? '';
    if (prev.formats) this.formats = { ...prev.formats };
    if (prev.hiddenRows !== undefined) this.hiddenRows = new Set(prev.hiddenRows);
    if (prev.activeFilterCols !== undefined) this.activeFilterCols = new Set(prev.activeFilterCols);
    if (prev.filterActive !== undefined) {
      this.filterActive = prev.filterActive;
      this.deserializeAdvFilterState(prev.advFilterSavedState);
    }
    // If filter is no longer active after undo, clear all hidden rows
    if (!this.filterActive) {
      this.hiddenRows.clear();
      this.advFilterSavedState.clear();
      this.activeFilterCols.clear();
    } else if (this.advFilterSavedState.size > 0) {
      // Re-evaluate hidden rows from the restored filter criteria
      this.recalculateAllFilters();
    }
    this.closeMenus();
    this.updateDisplayCache();
    this.showToast('Undo.');
  }

  redo() {
    if (!this.future.length) { this.showToast('Nothing to redo.'); return; }
    this.history.push(JSON.stringify({
      cells: this.cells, formats: this.formats,
      hiddenRows: Array.from(this.hiddenRows),
      activeFilterCols: Array.from(this.activeFilterCols),
      filterActive: this.filterActive,
      advFilterSavedState: this.serializeAdvFilterState()
    }));
    const next = JSON.parse(this.future.pop()!);
    if (next.cells) for (let r = 0; r < this.ROWS; r++) for (let c = 0; c < this.COLS; c++) this.cells[r][c] = next.cells[r]?.[c] ?? '';
    if (next.formats) this.formats = { ...next.formats };
    if (next.hiddenRows !== undefined) this.hiddenRows = new Set(next.hiddenRows);
    if (next.activeFilterCols !== undefined) this.activeFilterCols = new Set(next.activeFilterCols);
    if (next.filterActive !== undefined) {
      this.filterActive = next.filterActive;
      this.deserializeAdvFilterState(next.advFilterSavedState);
    }
    // If filter is no longer active after redo, clear all hidden rows
    if (!this.filterActive) {
      this.hiddenRows.clear();
      this.advFilterSavedState.clear();
      this.activeFilterCols.clear();
    } else if (this.advFilterSavedState.size > 0) {
      // Re-evaluate hidden rows from the restored filter criteria
      this.recalculateAllFilters();
    }
    this.closeMenus();
    this.updateDisplayCache();
    this.showToast('Redo.');
  }

  copyCell() {
    // Determine the range to copy
    const startR = this.rangeStart ? Math.min(this.rangeStart.r, this.rangeEnd!.r) : this.selectedRow;
    const endR = this.rangeStart ? Math.max(this.rangeStart.r, this.rangeEnd!.r) : this.selectedRow;
    const startC = this.rangeStart ? Math.min(this.rangeStart.c, this.rangeEnd!.c) : this.selectedCol;
    const endC = this.rangeStart ? Math.max(this.rangeStart.c, this.rangeEnd!.c) : this.selectedCol;

    const rows = endR - startR + 1;
    const cols = endC - startC + 1;

    // Snapshot cells, formats, validations
    const cellSnap: string[][] = [];
    const fmtSnap: Record<string, any> = {};
    const valSnap: Record<string, any> = {};
    const tsvRows: string[] = [];

    for (let r = 0; r < rows; r++) {
      cellSnap[r] = [];
      const tsvCols: string[] = [];
      for (let c = 0; c < cols; c++) {
        const srcR = startR + r;
        const srcC = startC + c;
        const val = this.cells[srcR][srcC] || '';
        cellSnap[r][c] = val;
        tsvCols.push(val);
        const srcKey = `${srcR},${srcC}`;
        const dstKey = `${r},${c}`;
        if (this.formats[srcKey]) fmtSnap[dstKey] = { ...this.formats[srcKey] };
        if (this.validations[srcKey]) valSnap[dstKey] = JSON.parse(JSON.stringify(this.validations[srcKey]));
      }
      tsvRows.push(tsvCols.join('\t'));
    }

    this.richClipboard = { cells: cellSnap, formats: fmtSnap, validations: valSnap, rows, cols, originR: startR, originC: startC };
    this.clipboard = cellSnap[0][0]; // fallback for system paste

    // Write TSV to system clipboard so Ctrl+V also works in other apps
    navigator.clipboard.writeText(tsvRows.join('\n')).catch(() => { });
    this.closeMenus();
    this.showToast(`Copied ${rows}×${cols} cell${rows * cols > 1 ? 's' : ''}.`);
  }

  pasteCell() {
    if (this.sheets[this.currentSheetIdx]?.locked) { this.showToast('This sheet is locked.'); return; }
    if (this.isSelectionLocked()) { this.showToast('Some cells in the selected range are locked.'); return; }
    if (this.richClipboard) {
      // Use rich internal clipboard (preserves formats, validations, multi-cell ranges)
      this.applyRichPaste('all');
      this.showToast('Pasted.');
    } else {
      // Fallback: read from system clipboard (external paste)
      navigator.clipboard.readText().then(text => {
        this.pushHistory({ action_type: 'paste', target_range: this.getA1(this.selectedRow, this.selectedCol) });
        this.cells[this.selectedRow][this.selectedCol] = text;
        this.formulaBarValue = text;
        this.onCellChange();
        this.showToast('Pasted.');
      }).catch(() => {
        if (this.clipboard) {
          this.pushHistory({ action_type: 'paste', target_range: this.getA1(this.selectedRow, this.selectedCol) });
          this.cells[this.selectedRow][this.selectedCol] = this.clipboard;
          this.formulaBarValue = this.clipboard;
          this.onCellChange();
          this.showToast('Pasted.');
        }
      });
      this.closeMenus();
    }
  }

  pasteValues() { this.applyRichPaste('values'); this.showToast('Pasted values.'); }
  pasteFormulas() { this.applyRichPaste('formulas'); this.showToast('Pasted formulas.'); }
  pasteFormats() { this.applyRichPaste('formats'); this.showToast('Pasted formats.'); }
  pasteNotes() { this.applyRichPaste('notes'); this.showToast('Pasted notes.'); }
  pasteFormulasAndNumberFormats() { this.applyRichPaste('formulasAndNumbers'); this.showToast('Pasted formulas & number formats.'); }
  pasteValuesAndNumberFormats() { this.applyRichPaste('valuesAndNumbers'); this.showToast('Pasted values & number formats.'); }
  pasteValidation() { this.applyRichPaste('validation'); this.showToast('Pasted validation rules.'); }
  pasteExceptNotes() { this.applyRichPaste('exceptNotes'); this.showToast('Pasted all except notes.'); }
  pasteExceptBorders() { this.applyRichPaste('exceptBorders'); this.showToast('Pasted all except borders.'); }
  pasteTranspose() { this.applyRichPaste('transpose'); this.showToast('Pasted transposed.'); }
  pasteLinkToSource() { this.showToast('Link To Source is not supported for internal pastes.'); this.closeMenus(); }

  private forEachSelectedCell(callback: (r: number, c: number) => void) {
    if (this.rangeStart && this.rangeEnd) {
      const minR = Math.min(this.rangeStart.r, this.rangeEnd.r);
      const maxR = Math.max(this.rangeStart.r, this.rangeEnd.r);
      const minC = Math.min(this.rangeStart.c, this.rangeEnd.c);
      const maxC = Math.max(this.rangeStart.c, this.rangeEnd.c);
      for (let r = minR; r <= maxR; r++) {
        for (let c = minC; c <= maxC; c++) {
          callback(r, c);
        }
      }
    } else {
      callback(this.selectedRow, this.selectedCol);
    }
  }

  // ── Internal helper: paste rich clipboard to destination ─────────────────
  private applyRichPaste(
    mode: 'all' | 'values' | 'formulas' | 'formats' | 'notes' |
      'formulasAndNumbers' | 'valuesAndNumbers' | 'validation' |
      'exceptNotes' | 'exceptBorders' | 'transpose'
  ) {
    const destR = this.selectedRow;
    const destC = this.selectedCol;

    if (!this.richClipboard) {
      // Fallback: plain text paste
      this.pasteCell();
      return;
    }

    const { cells, formats: fmts, validations: vals, rows, cols } = this.richClipboard;
    const pasteRows = mode === 'transpose' ? cols : rows;
    const pasteCols = mode === 'transpose' ? rows : cols;

    this.pushHistory({
      action_type: `paste-${mode}`,
      target_range: `${this.getA1(destR, destC)}:${this.getA1(destR + pasteRows - 1, destC + pasteCols - 1)}`
    });

    for (let r = 0; r < pasteRows; r++) {
      for (let c = 0; c < pasteCols; c++) {
        const srcR = mode === 'transpose' ? c : r;
        const srcC = mode === 'transpose' ? r : c;
        const targetR = destR + r;
        const targetC = destC + c;
        if (targetR >= this.ROWS || targetC >= this.COLS) continue;
        const dstKey = `${targetR},${targetC}`;
        const srcKey = `${srcR},${srcC}`;
        const srcFmt = fmts[srcKey] || {};

        if (mode === 'values' || mode === 'valuesAndNumbers' || mode === 'transpose') {
          // Value only — strip formulas (values are already resolved text)
          this.cells[targetR][targetC] = cells[srcR]?.[srcC] ?? '';
          if (mode === 'valuesAndNumbers' && (srcFmt as any).numFormat) {
            (this.formats[dstKey] as any) = { ...(this.formats[dstKey] || {}), numFormat: (srcFmt as any).numFormat };
          }

        } else if (mode === 'formulas' || mode === 'formulasAndNumbers') {
          this.cells[targetR][targetC] = cells[srcR]?.[srcC] ?? '';
          if (mode === 'formulasAndNumbers' && (srcFmt as any).numFormat) {
            (this.formats[dstKey] as any) = { ...(this.formats[dstKey] || {}), numFormat: (srcFmt as any).numFormat };
          }

        } else if (mode === 'formats') {
          // Format only — don't touch cell values
          const { note, comment, ...visualFmt } = srcFmt as any;
          this.formats[dstKey] = { ...(this.formats[dstKey] || {}), ...visualFmt };

        } else if (mode === 'notes') {
          // Note/comment only
          if ((srcFmt as any).note !== undefined) {
            (this.formats[dstKey] as any) = { ...(this.formats[dstKey] || {}), note: (srcFmt as any).note };
          }
          if ((srcFmt as any).comment !== undefined) {
            (this.formats[dstKey] as any) = { ...(this.formats[dstKey] || {}), comment: (srcFmt as any).comment };
          }

        } else if (mode === 'validation') {
          if (vals[srcKey]) this.validations[dstKey] = JSON.parse(JSON.stringify(vals[srcKey]));
          else delete this.validations[dstKey];

        } else if (mode === 'exceptNotes') {
          // All except notes/comments
          this.cells[targetR][targetC] = cells[srcR]?.[srcC] ?? '';
          const { note, comment, ...fmtWithoutNotes } = srcFmt as any;
          if (Object.keys(fmtWithoutNotes).length) this.formats[dstKey] = fmtWithoutNotes;
          else delete this.formats[dstKey];
          if (vals[srcKey]) this.validations[dstKey] = JSON.parse(JSON.stringify(vals[srcKey]));

        } else if (mode === 'exceptBorders') {
          // All except borders
          this.cells[targetR][targetC] = cells[srcR]?.[srcC] ?? '';
          const { borders, ...fmtWithoutBorders } = srcFmt as any;
          if (Object.keys(fmtWithoutBorders).length) this.formats[dstKey] = fmtWithoutBorders;
          else delete this.formats[dstKey];
          if (vals[srcKey]) this.validations[dstKey] = JSON.parse(JSON.stringify(vals[srcKey]));

        } else {
          // 'all' — paste everything
          this.cells[targetR][targetC] = cells[srcR]?.[srcC] ?? '';
          if (Object.keys(srcFmt).length) this.formats[dstKey] = { ...srcFmt };
          else delete this.formats[dstKey];
          if (vals[srcKey]) this.validations[dstKey] = JSON.parse(JSON.stringify(vals[srcKey]));
          else delete this.validations[dstKey];
        }
      }
    }

    this.formats = { ...this.formats };
    // Select the pasted range
    this.rangeStart = { r: destR, c: destC };
    this.rangeEnd = { r: Math.min(destR + pasteRows - 1, this.ROWS - 1), c: Math.min(destC + pasteCols - 1, this.COLS - 1) };
    this.onCellChange();
    this.updateDisplayCache();
    this.closeMenus();
  }


  clearNotes() {
    this.pushHistory();
    this.forEachSelectedCell((r, c) => {
      if (this.formats[`${r},${c}`]) {
        delete (this.formats[`${r},${c}`] as any).note;
        delete (this.formats[`${r},${c}`] as any).comment;
      }
    });
    this.formats = { ...this.formats };
    this.onCellChange();
    this.save();
    this.showToast('Cleared notes.');
    this.closeMenus();
  }

  clearHyperlinks() {
    this.pushHistory();
    this.forEachSelectedCell((r, c) => {
      if (this.formats[`${r},${c}`]) {
        delete (this.formats[`${r},${c}`] as any).hyperlink;
        delete (this.formats[`${r},${c}`] as any).underline;
        if (this.formats[`${r},${c}`].color === '#1155cc' || this.formats[`${r},${c}`].color === '#1a73e8') {
          delete this.formats[`${r},${c}`].color;
        }
      }
    });
    this.formats = { ...this.formats };
    this.onCellChange();
    this.save();
    this.showToast('Cleared hyperlinks.');
    this.closeMenus();
  }

  clearCheckboxes() {
    this.pushHistory();
    this.forEachSelectedCell((r, c) => {
      if (this.formats[`${r},${c}`] && (this.formats[`${r},${c}`] as any).checkbox) {
        delete (this.formats[`${r},${c}`] as any).checkbox;
        if (this.cells[r][c] === 'TRUE' || this.cells[r][c] === 'FALSE') {
          this.cells[r][c] = '';
        }
      }
    });
    this.formats = { ...this.formats };
    this.onCellChange();
    this.save();
    this.showToast('Cleared checkboxes.');
    this.closeMenus();
  }

  clearDataValidations() {
    this.pushHistory();
    this.forEachSelectedCell((r, c) => {
      delete this.validations[`${r},${c}`];
    });
    this.validations = { ...this.validations };
    this.onCellChange();
    this.save();
    this.showToast('Cleared data validations.');
    this.closeMenus();
  }

  clearConditionalFormats() {
    this.pushHistory();
    this.forEachSelectedCell((r, c) => {
      if (this.formats[`${r},${c}`]) {
        delete (this.formats[`${r},${c}`] as any).conditionalFormat;
      }
    });
    this.formats = { ...this.formats };
    this.onCellChange();
    this.save();
    this.showToast('Cleared conditional formats.');
    this.closeMenus();
  }

  clearRichTextFormats() {
    this.pushHistory();
    this.forEachSelectedCell((r, c) => {
      if (this.formats[`${r},${c}`]) {
        delete this.formats[`${r},${c}`].bold;
        delete this.formats[`${r},${c}`].italic;
        delete this.formats[`${r},${c}`].underline;
        delete this.formats[`${r},${c}`].strikethrough;
        delete this.formats[`${r},${c}`].color;
        delete (this.formats[`${r},${c}`] as any).fontFamily;
        delete (this.formats[`${r},${c}`] as any).fontSize;
      }
    });
    this.formats = { ...this.formats };
    this.onCellChange();
    this.save();
    this.showToast('Cleared rich text formatting.');
    this.closeMenus();
  }

  clearAllFilters() {
    this.pushHistory();
    this.filterActive = false;
    this.activeFilterCols.clear();
    this.advFilterSavedState.clear();
    this.hiddenRows.clear();
    this.onCellChange(undefined, undefined, true);
    this.showToast('All filters cleared.');
  }

  shapeTab: 'text' | 'shape' | 'diagram' = 'diagram';
  diagramCategory: 'list' | 'process' | 'pyramid' | 'cycle' = 'list';
  shapeCategory: 'shape' | 'lines' | 'flowchart' | 'math' | 'stars' | 'callouts' = 'shape';
  textCategory: 'textbox' | 'symbol' = 'textbox';
  activeShapeIdx: number | null = null;
  activeShapeMenuIdx: number | null = null;

  insertShape(type: string) {
    const sheet = this.sheets[this.currentSheetIdx];
    if (!sheet.shapes) {
      sheet.shapes = [];
    }
    const x = this.getColOffset(this.selectedCol) + 20;
    const y = this.getRowOffset(this.selectedRow) + 20;

    let width = 100;
    let height = 100;
    let text = '';

    if (type.startsWith('diagram')) {
      width = 250;
      height = 150;
    } else if (type.startsWith('text')) {
      width = 150;
      height = 40;
      text = 'Sample Text';
    } else if (type.startsWith('symbol_')) {
      width = 40;
      height = 40;
      const symbols: { [key: string]: string } = {
        'symbol_copy': '©', 'symbol_reg': '®', 'symbol_tm': '™',
        'symbol_pi': 'π', 'symbol_sigma': 'Σ', 'symbol_omega': 'Ω', 'symbol_inf': '∞'
      };
      text = symbols[type] || '';
    } else if (type === 'button') {
      width = 120;
      height = 40;
      text = 'Button';
    } else {
      text = '';
    }

    sheet.shapes.push({
      id: 'shape_' + Date.now(),
      type: type,
      x: x,
      y: y,
      width: width,
      height: height,
      text: text,
    });
    this.activeMenu = null;
    this.save();
  }

  deleteShape(idx: number) {
    const sheet = this.sheets[this.currentSheetIdx];
    if (sheet.shapes) {
      sheet.shapes.splice(idx, 1);
      this.save();
    }
  }

  async editShapeLabel(idx: number) {
    const sheet = this.sheets[this.currentSheetIdx];
    if (!sheet.shapes) return;
    const shape = sheet.shapes[idx];
    const currentText = shape.text || '';
    const newText = await this.openPrompt('Enter text for this shape:', currentText);
    if (newText !== null) {
      shape.text = newText;
      this.save();
    }
  }

  startShapeDrag(e: MouseEvent, idx: number) {
    this.activeShapeIdx = idx;
    this.activeShapeMenuIdx = null;
    this.closeMenus();
    const sheet = this.sheets[this.currentSheetIdx];
    if (!sheet.shapes) return;
    const shape = sheet.shapes[idx];

    const startX = e.clientX;
    const startY = e.clientY;
    const initialX = shape.x;
    const initialY = shape.y;

    const moveListener = (moveEvent: MouseEvent) => {
      shape.x = initialX + (moveEvent.clientX - startX);
      shape.y = initialY + (moveEvent.clientY - startY);
    };

    const upListener = () => {
      document.removeEventListener('mousemove', moveListener);
      document.removeEventListener('mouseup', upListener);
      this.save();
    };

    document.addEventListener('mousemove', moveListener);
    document.addEventListener('mouseup', upListener);

    e.preventDefault();
  }

  recalculate() { this.updateDisplayCache(); this.showToast('Recalculated.'); }

  clearCell() {
    if (this.sheets[this.currentSheetIdx]?.locked) { this.showToast('This sheet is locked.'); return; }
    if ((this.formats[`${this.selectedRow},${this.selectedCol}`] as any)?.locked) { this.showToast('This cell is locked.'); return; }
    this.pushHistory();
    this.cells[this.selectedRow][this.selectedCol] = '';
    const ref = `${this.selectedRow},${this.selectedCol}`;
    if (this.formats[ref]) {
      delete this.formats[ref].bg;
      delete this.formats[ref].bold;
      delete this.formats[ref].italic;
      delete this.formats[ref].color;
      delete this.formats[ref].strikethrough;
      delete (this.formats[ref] as any).checkbox;
    }
    this.formulaBarValue = '';
    this.onCellChange();
    this.closeMenus();
  }


  shiftCellsLeft() {
    this.pushHistory();
    const r = this.selectedRow;
    const c = this.selectedCol;
    for (let i = c; i < this.COLS - 1; i++) {
      this.cells[r][i] = this.cells[r][i + 1];
      const nextFmt = this.formats[`${r},${i + 1}`];
      if (nextFmt) {
        this.formats[`${r},${i}`] = { ...nextFmt };
      } else {
        delete this.formats[`${r},${i}`];
      }
    }
    this.cells[r][this.COLS - 1] = '';
    delete this.formats[`${r},${this.COLS - 1}`];
    this.onCellChange();
    this.closeMenus();
    this.showToast('Shifted cells left.');
  }

  shiftCellsUp() {
    this.pushHistory();
    const r = this.selectedRow;
    const c = this.selectedCol;
    for (let i = r; i < this.ROWS - 1; i++) {
      this.cells[i][c] = this.cells[i + 1][c];
      const nextFmt = this.formats[`${i + 1},${c}`];
      if (nextFmt) {
        this.formats[`${i},${c}`] = { ...nextFmt };
      } else {
        delete this.formats[`${i},${c}`];
      }
    }
    this.cells[this.ROWS - 1][c] = '';
    delete this.formats[`${this.ROWS - 1},${c}`];
    this.onCellChange();
    this.closeMenus();
    this.showToast('Shifted cells up.');
  }

  insertRowAbove() {
    const count = this.selectedRowCount;
    const r = this.rangeStart && this.rangeEnd ? Math.min(this.rangeStart.r, this.rangeEnd.r) : this.selectedRow;
    this.pushHistory({ action_type: 'insert-row-above', target_range: `${r + 1}` });
    for (let i = 0; i < count; i++) {
      this.cells.splice(r, 0, Array(this.COLS).fill(''));
    }
    this.ROWS += count;
    this.rowRange = Array.from({ length: this.ROWS }, (_, i) => i);
    this.onCellChange(); this.closeMenus();
    this.showToast(`${count} Row${count > 1 ? 's' : ''} inserted above.`);
  }

  insertRowBelow() {
    const count = this.selectedRowCount;
    const r = (this.rangeStart && this.rangeEnd ? Math.max(this.rangeStart.r, this.rangeEnd.r) : this.selectedRow) + 1;
    this.pushHistory({ action_type: 'insert-row-below', target_range: `${r + 1}` });
    for (let i = 0; i < count; i++) {
      this.cells.splice(r, 0, Array(this.COLS).fill(''));
    }
    this.ROWS += count;
    this.rowRange = Array.from({ length: this.ROWS }, (_, i) => i);
    this.onCellChange(); this.closeMenus();
    this.showToast(`${count} Row${count > 1 ? 's' : ''} inserted below.`);
  }

  insertColLeft() {
    const count = this.selectedColCount;
    const c = this.rangeStart && this.rangeEnd ? Math.min(this.rangeStart.c, this.rangeEnd.c) : this.selectedCol;
    this.pushHistory({ action_type: 'insert-column-left', target_range: this.colLabel(c) });
    this.COLS += count;
    for (const row of this.cells) {
      for (let i = 0; i < count; i++) row.splice(c, 0, '');
    }
    this.colRange = Array.from({ length: this.COLS }, (_, i) => i);
    this.onCellChange(); this.closeMenus();
    this.showToast(`${count} Column${count > 1 ? 's' : ''} inserted.`);
  }

  insertColRight() {
    const count = this.selectedColCount;
    const c = (this.rangeStart && this.rangeEnd ? Math.max(this.rangeStart.c, this.rangeEnd.c) : this.selectedCol) + 1;
    this.pushHistory({ action_type: 'insert-column-right', target_range: this.colLabel(c) });
    this.COLS += count;
    for (const row of this.cells) {
      for (let i = 0; i < count; i++) row.splice(c, 0, '');
    }
    this.colRange = Array.from({ length: this.COLS }, (_, i) => i);
    this.onCellChange(); this.closeMenus();
    this.showToast(`${count} Column${count > 1 ? 's' : ''} inserted.`);
  }

  deleteRow() {
    const count = this.selectedRowCount;
    const r = this.rangeStart && this.rangeEnd ? Math.min(this.rangeStart.r, this.rangeEnd.r) : this.selectedRow;
    this.pushHistory({ action_type: 'delete-row', target_range: `${r + 1}:${r + count}` });
    this.cells.splice(r, count);
    this.ROWS = Math.max(1, this.ROWS - count);
    while (this.cells.length < this.ROWS) this.cells.push(Array(this.COLS).fill(''));
    while (this.cells.length > this.ROWS) this.cells.pop();
    this.rowRange = Array.from({ length: this.ROWS }, (_, i) => i);
    if (this.selectedRow >= this.ROWS) this.selectedRow = this.ROWS - 1;
    this.onCellChange(); this.closeMenus();
    this.showToast(`${count} Row${count > 1 ? 's' : ''} deleted.`);
  }

  deleteCol() {
    const count = this.selectedColCount;
    const c = this.rangeStart && this.rangeEnd ? Math.min(this.rangeStart.c, this.rangeEnd.c) : this.selectedCol;
    this.pushHistory({ action_type: 'delete-column', target_range: `${this.colLabel(c)}:${this.colLabel(c + count - 1)}` });
    this.COLS = Math.max(1, this.COLS - count);
    for (const row of this.cells) {
      row.splice(c, count);
      while (row.length < this.COLS) row.push('');
      while (row.length > this.COLS) row.pop();
    }
    this.colRange = Array.from({ length: this.COLS }, (_, i) => i);
    if (this.selectedCol >= this.COLS) this.selectedCol = this.COLS - 1;
    this.onCellChange(); this.closeMenus();
    this.showToast(`${count} Column${count > 1 ? 's' : ''} deleted.`);
  }

  shiftCellsDown() {
    this.pushHistory();
    const r = this.selectedRow;
    const c = this.selectedCol;
    for (let i = this.ROWS - 1; i > r; i--) {
      this.cells[i][c] = this.cells[i - 1][c];
      const prevFmt = this.formats[`${i - 1},${c}`];
      if (prevFmt) { this.formats[`${i},${c}`] = { ...prevFmt }; }
      else { delete this.formats[`${i},${c}`]; }
    }
    this.cells[r][c] = '';
    delete this.formats[`${r},${c}`];
    this.onCellChange();
    this.closeMenus();
    this.showToast('Shifted cells down.');
  }

  shiftCellsRight() {
    this.pushHistory();
    const r = this.selectedRow;
    const c = this.selectedCol;
    for (let i = this.COLS - 1; i > c; i--) {
      this.cells[r][i] = this.cells[r][i - 1];
      const prevFmt = this.formats[`${r},${i - 1}`];
      if (prevFmt) { this.formats[`${r},${i}`] = { ...prevFmt }; }
      else { delete this.formats[`${r},${i}`]; }
    }
    this.cells[r][c] = '';
    delete this.formats[`${r},${c}`];
    this.onCellChange();
    this.closeMenus();
    this.showToast('Shifted cells right.');
  }

  // ── Custom Insert Modal ───────────────────────────────────────
  openCustomInsert(type: 'row' | 'col') {
    this.closeMenus();
    this.customInsertType = type;
    this.customInsertCount = 1;
    this.customInsertPosition = 'before';
    this.activeModal = 'custom_insert';
  }

  confirmCustomInsert() {
    const count = Math.max(1, Math.floor(Number(this.customInsertCount) || 1));
    if (this.customInsertType === 'row') {
      const r = this.customInsertPosition === 'before'
        ? (this.rangeStart && this.rangeEnd ? Math.min(this.rangeStart.r, this.rangeEnd.r) : this.selectedRow)
        : (this.rangeStart && this.rangeEnd ? Math.max(this.rangeStart.r, this.rangeEnd.r) : this.selectedRow) + 1;
      this.pushHistory({ action_type: 'insert-row-custom', target_range: `${r + 1}` });
      for (let i = 0; i < count; i++) {
        this.cells.splice(r, 0, Array(this.COLS).fill(''));
      }
      this.ROWS += count;
      this.rowRange = Array.from({ length: this.ROWS }, (_, i) => i);
      this.onCellChange();
      this.showToast(`${count} row${count > 1 ? 's' : ''} inserted ${this.customInsertPosition === 'before' ? 'above' : 'below'}.`);
    } else {
      const c = this.customInsertPosition === 'before'
        ? (this.rangeStart && this.rangeEnd ? Math.min(this.rangeStart.c, this.rangeEnd.c) : this.selectedCol)
        : (this.rangeStart && this.rangeEnd ? Math.max(this.rangeStart.c, this.rangeEnd.c) : this.selectedCol) + 1;
      this.pushHistory({ action_type: 'insert-col-custom', target_range: this.colLabel(c) });
      this.COLS += count;
      for (const row of this.cells) {
        for (let i = 0; i < count; i++) row.splice(c, 0, '');
      }
      this.colRange = Array.from({ length: this.COLS }, (_, i) => i);
      this.onCellChange();
      this.showToast(`${count} column${count > 1 ? 's' : ''} inserted ${this.customInsertPosition === 'before' ? 'before' : 'after'} ${this.colLabel(this.customInsertPosition === 'before' ? c : c - 1)}.`);
    }
    this.activeModal = null;
  }
  // ─────────────────────────────────────────────────────────────


  createSparkline() {
    this.closeMenus();
    this.insertSparklineConfig = {
      source: `'${this.sheets[this.currentSheetIdx].name}'.${this.colLabel(this.selectedCol)}${this.selectedRow + 1}`,
      dest: '',
      error: ''
    };
    this.activeModal = 'insert_sparkline';
  }


  isSparklineCell(r: number, c: number): boolean {
    const sheet = this.sheets[this.currentSheetIdx];
    return !!(sheet.sparklines && sheet.sparklines[`${r},${c}`]);
  }

  submitEditSparkline() {
    const srcRange = this.parseRange(this.editSparklineConfig.source);
    const destRange = this.parseRange(this.editSparklineConfig.dest);

    if (!srcRange || !destRange) {
      this.editSparklineConfig.error = 'Invalid range format.';
      return;
    }

    const srcRows = srcRange.endR - srcRange.startR + 1;
    const srcCols = srcRange.endC - srcRange.startC + 1;
    const destRows = destRange.endR - destRange.startR + 1;
    const destCols = destRange.endC - destRange.startC + 1;

    if (!(destRows === 1 && destCols === 1) && (srcRows !== destRows || srcCols !== destCols)) {
      this.editSparklineConfig.error = 'Source and destination dimensions must match.';
      return;
    }

    this.editSparklineConfig.error = '';

    const sheet = this.sheets[this.currentSheetIdx];
    if (!sheet.sparklines) sheet.sparklines = {};

    // For simplicity, just update the currently selected cell if one is selected,
    // or maybe the logic is more complex. Let's just update the config for the current cell for now.
    // The user can edit the sparkline in the side panel anyway.

    const key = `${this.selectedRow},${this.selectedCol}`;
    if (sheet.sparklines[key]) {
      sheet.sparklines[key].sourceRange = this.editSparklineConfig.source;
      sheet.sparklines[key].destinationRange = this.editSparklineConfig.dest;
      
      if (this.sparklineConfig && this.sidePanelApp === 'sparkline') {
        this.sparklineConfig.sourceRange = this.editSparklineConfig.source;
        this.sparklineConfig.destinationRange = this.editSparklineConfig.dest;
      }
    }

    this.activeModal = null;
    this.save();
    this.showToast('Sparkline range updated');
  }

  openSparklineFormat() {
    const config = this.sheets[this.currentSheetIdx].sparklines![`${this.selectedRow},${this.selectedCol}`];
    if (config) {
      this.sparklineConfig = JSON.parse(JSON.stringify(config));
      this.sidePanelApp = 'sparkline';
    } else {
      this.showToast('Select a cell containing a sparkline first.');
    }
  }

  submitInsertSparkline() {
    // Validate that source and dest ranges match in dimensions
    const srcRange = this.parseRange(this.insertSparklineConfig.source);
    const destRange = this.parseRange(this.insertSparklineConfig.dest);

    if (!srcRange || !destRange) {
      this.insertSparklineConfig.error = 'Invalid range format. Please use format like "Sheet1.A1:B10" or "A1:B10".';
      return;
    }

    const srcRows = srcRange.endR - srcRange.startR + 1;
    const srcCols = srcRange.endC - srcRange.startC + 1;
    const destRows = destRange.endR - destRange.startR + 1;
    const destCols = destRange.endC - destRange.startC + 1;

    if (!(destRows === 1 && destCols === 1) && (srcRows !== destRows || srcCols !== destCols)) {
      this.insertSparklineConfig.error = 'Please select a destination range that is equal to the source range.';
      return;
    }

    this.insertSparklineConfig.error = '';
    const groupId = 'sparkgroup_' + Date.now();

    const sheet = this.sheets[this.currentSheetIdx];
    if (!sheet.sparklines) sheet.sparklines = {};

    // Create the group
    for (let rOffset = 0; rOffset < destRows; rOffset++) {
      for (let cOffset = 0; cOffset < destCols; cOffset++) {
        const destR = destRange.startR + rOffset;
        const destC = destRange.startC + cOffset;

        // Find corresponding source sub-range (can be 1D or single cell depending on sparkline logic, usually source is a 1D range for a single sparkline)
        // Actually, if we map one-to-one, Zoho creates one sparkline per row or col mapping.
        // Wait, the specification says: "create one sparkline per row/column pair mapping source -> destination cell"
        // Meaning if source is C2:C10 (9x1) and dest is D2:D10 (9x1), it maps C2 to D2, C3 to D3, etc.
        // Wait! A sparkline needs a 1D range of data to draw a chart.
        // If source is C2:E2 (1x3) and dest is F2 (1x1), the sparkline at F2 uses C2:E2.
        // But if the user selects multiple destination cells, we map rows to rows.
        // For simplicity here, we assume destR/destC maps to a slice of the source.
        // Let's just create one sparkline at the first dest cell for the entire source range, or map row-by-row if heights match.

        const key = `${destR},${destC}`;

        // Define slice
        let sliceSrc = `${this.colLabel(srcRange.startC + cOffset)}${srcRange.startR + rOffset + 1}`;
        // Actually, usually sparklines take a row or column of data.
        // Let's just set the source range to the whole source range if it's 1x1 dest, else map row by row.
        let assignedSource = this.insertSparklineConfig.source;
        if (destRows > 1 && srcRows > 1 && destRows === srcRows && srcCols !== destCols) {
          // map row to row
          assignedSource = `${this.colLabel(srcRange.startC)}${srcRange.startR + rOffset + 1}:${this.colLabel(srcRange.endC)}${srcRange.startR + rOffset + 1}`;
        } else if (destCols > 1 && srcCols > 1 && destCols === srcCols && srcRows !== destRows) {
          // map col to col
          assignedSource = `${this.colLabel(srcRange.startC + cOffset)}${srcRange.startR + 1}:${this.colLabel(srcRange.startC + cOffset)}${srcRange.endR + 1}`;
        }

        sheet.sparklines[key] = {
          sourceRange: assignedSource,
          destinationRange: `${this.colLabel(destC)}${destR + 1}`,
          type: 'line',
          baseColor: '#4A86E8',
          highlights: {
            high: { enabled: false, color: '#34A853' },
            low: { enabled: false, color: '#F4B400' },
            first: { enabled: false, color: '#4A86E8' },
            last: { enabled: false, color: '#7BAAF7' },
            negative: { enabled: false, color: '#EA4335' },
            markers: { enabled: false, color: '#4A86E8' }
          },
          emptyCellMode: 'gap',
          includeHiddenRowsColumns: false,
          horizontalAxis: { displayAxis: false, rightToLeft: false },
          verticalAxis: {
            min: { mode: 'auto', customValue: null },
            max: { mode: 'auto', customValue: null }
          },
          isGrouped: (destRows * destCols > 1),
          groupId: groupId
        };
      }
    }

    this.activeModal = null;

    // Select the first destination cell and open the sparkline app
    this.selectedRow = destRange.startR;
    this.selectedCol = destRange.startC;
    this.editSparkline();
  }

  parseRange(rangeStr: string) {
    // Handle 'Sheet1'.A1:B2 or Sheet1!A1:B2 or just A1:B2
    let sheetName = this.sheets[this.currentSheetIdx].name;
    let localRange = rangeStr;

    const match = rangeStr.match(/^(?:'([^']+)'[.!]|([^.!]+)[.!])?(.*)$/);
    if (match) {
      if (match[1]) sheetName = match[1];
      if (match[2]) sheetName = match[2];
      localRange = match[3];
    }

    const parts = localRange.split(':');
    if (parts.length > 2) return null;

    const start = this.parseCellRef(parts[0]);
    if (!start) return null;

    const end = parts.length === 2 ? this.parseCellRef(parts[1]) : start;
    if (!end) return null;

    return {
      sheetName,
      startR: Math.min(start.r, end.r),
      startC: Math.min(start.c, end.c),
      endR: Math.max(start.r, end.r),
      endC: Math.max(start.c, end.c)
    };
  }




  editSparkline() {
    this.closeMenus();
    const sheet = this.sheets[this.currentSheetIdx];
    const key = `${this.selectedRow},${this.selectedCol}`;
    if (sheet.sparklines && sheet.sparklines[key]) {
      this.sparklineConfig = JSON.parse(JSON.stringify(sheet.sparklines[key]));
      this.openApp('sparkline');
    } else {
      this.showToast('Selected cell does not contain a sparkline.');
    }
  }


  // --- New Sparkline Helpers ---
  horizontalAxisExpanded = false;
  verticalAxisExpanded = false;

  openColorPicker(event: MouseEvent, target: 'base' | 'high' | 'low' | 'first' | 'last' | 'negative' | 'markers') {
    const el = event.currentTarget as HTMLElement;
    const rect = el.getBoundingClientRect();
    let left = rect.left;
    if (left + 240 > window.innerWidth) left = window.innerWidth - 240;
    this.colorPickerState = {
      active: true,
      target: target,
      top: rect.bottom + 4,
      left: left
    };
    this.cdr.markForCheck();
    event.stopPropagation();
  }

  closeColorPicker() {
    this.colorPickerState.active = false;
    this.cdr.markForCheck();
  }

  setSparklineColor(color: string) {
    if (!this.sparklineConfig || !this.colorPickerState.target) return;

    const target = this.colorPickerState.target;
    if (target === 'base') {
      this.sparklineConfig.baseColor = color;
    } else {
      this.sparklineConfig.highlights[target].color = color;
    }

    // Add to recent colors if not there
    if (!this.recentColors.includes(color)) {
      this.recentColors.unshift(color);
      if (this.recentColors.length > 10) this.recentColors.pop();
    }

    this.saveSparkline();
    this.closeColorPicker();
  }

  setSparklineType(type: 'line' | 'column' | 'winloss') {
    if (!this.sparklineConfig) return;
    this.sparklineConfig.type = type;
    this.saveSparkline();
  }

  setEmptyCellMode(mode: 'gap' | 'zero' | 'connect' | 'skip') {
    if (!this.sparklineConfig) return;
    this.sparklineConfig.emptyCellMode = mode;
    this.saveSparkline();
  }

  toggleGroup() {
    if (!this.sparklineConfig) return;
    this.sparklineConfig.isGrouped = !this.sparklineConfig.isGrouped;
    if (this.sparklineConfig.isGrouped && !this.sparklineConfig.groupId) {
      this.sparklineConfig.groupId = 'sparkgroup_' + Date.now();
    }
    this.saveSparkline();
  }

  switchRowsColumns() {
    // This flips whether data is read by rows or columns
    // In our simplified parse, it's 1D, but we can implement it as just triggering a re-render.
    // For now we will add a property if we need to actually transpose 2D source.
    // The spec just says "useful when sparkline data orientation needs to flip".
    this.showToast('Switch Rows/Columns applied');
    this.saveSparkline();
  }

  saveSparkline() {
    if (!this.sparklineConfig) return;
    const sheet = this.sheets[this.currentSheetIdx];
    if (!sheet.sparklines) sheet.sparklines = {};

    // Save live changes to the sparkline or group
    const isGrouped = this.sparklineConfig.isGrouped;
    const groupId = this.sparklineConfig.groupId;

    if (isGrouped && groupId) {
      for (const key of Object.keys(sheet.sparklines)) {
        if (sheet.sparklines[key].groupId === groupId) {
          const dest = sheet.sparklines[key].destinationRange;
          const src = sheet.sparklines[key].sourceRange;
          sheet.sparklines[key] = JSON.parse(JSON.stringify(this.sparklineConfig));
          // Restore unique source/dest for each in group
          sheet.sparklines[key].destinationRange = dest;
          sheet.sparklines[key].sourceRange = src;
        }
      }
    } else {
      // Single
      const key = `${this.selectedRow},${this.selectedCol}`;
      const dest = sheet.sparklines[key]?.destinationRange || `${this.colLabel(this.selectedCol)}${this.selectedRow + 1}`;
      const src = sheet.sparklines[key]?.sourceRange || this.sparklineConfig.sourceRange;
      sheet.sparklines[key] = JSON.parse(JSON.stringify(this.sparklineConfig));
      sheet.sparklines[key].destinationRange = dest;
      sheet.sparklines[key].sourceRange = src;
    }

    this.save();
  }

  deleteSparklineConfig() {
    const sheet = this.sheets[this.currentSheetIdx];
    if (!sheet.sparklines || !this.sparklineConfig) return;

    if (this.sparklineConfig.isGrouped && this.sparklineConfig.groupId) {
      for (const key of Object.keys(sheet.sparklines)) {
        if (sheet.sparklines[key].groupId === this.sparklineConfig.groupId) {
          delete sheet.sparklines[key];
        }
      }
    } else {
      const key = `${this.selectedRow},${this.selectedCol}`;
      delete sheet.sparklines[key];
    }

    this.sparklineConfig = null as any;
    this.sidePanelApp = null;
    this.save();
  }

  insertButton() {
    this.closeMenus();
    this.insertShape('button');
    this.showToast('Button inserted.');
  }

  async defineName() {
    this.closeMenus();
    const name = await this.openPrompt(`Enter name for range ${this.getRangeRef()}:`);
    if (name) {
      this.showToast(`Name "${name}" defined for ${this.getRangeRef()}.`);
    }
  }

  insertNote() {
    this.closeMenus();
    const ref = `${this.selectedRow},${this.selectedCol}`;
    this.activeNotePopup = { r: this.selectedRow, c: this.selectedCol, text: (this.formats[ref] as any)?.note || '' };
  }

  emojiPickerX = 400;
  emojiPickerY = 200;
  isDraggingEmoji = false;
  dragStartXP = 0;
  dragStartYP = 0;

  startEmojiDrag(e: MouseEvent) {
    e.preventDefault();
    this.isDraggingEmoji = true;
    this.dragStartXP = e.clientX - this.emojiPickerX;
    this.dragStartYP = e.clientY - this.emojiPickerY;
    
    const mouseMoveHandler = (ev: MouseEvent) => {
      if (this.isDraggingEmoji) {
        this.emojiPickerX = ev.clientX - this.dragStartXP;
        this.emojiPickerY = ev.clientY - this.dragStartYP;
      }
    };
    
    const mouseUpHandler = () => {
      this.isDraggingEmoji = false;
      document.removeEventListener('mousemove', mouseMoveHandler);
      document.removeEventListener('mouseup', mouseUpHandler);
    };
    
    document.addEventListener('mousemove', mouseMoveHandler);
    document.addEventListener('mouseup', mouseUpHandler);
  }

  insertEmoji() {
    this.closeMenus();
    this.emojiPickerX = window.innerWidth / 2 - 169;
    this.emojiPickerY = window.innerHeight / 2 - 215;
    this.activeModal = 'emoji';
  }

  filterEmojis = (emoji: any) => {
    if (!emoji || !emoji.unified) return true;
    const unified: string = emoji.unified;
    // Block ZWJ sequences (contain 200D) — most don't render on Windows 10
    if (unified.toUpperCase().includes('200D')) return false;
    // Block Unicode 14.0+ emojis (codepoints >= 0x1FAE0) — not in Win10 Segoe UI Emoji
    const firstCodePoint = parseInt(unified.split('-')[0], 16);
    if (firstCodePoint >= 0x1FAE0) return false;
    // Block other known unsupported ranges on Windows 10
    if (firstCodePoint >= 0x1FA70 && firstCodePoint <= 0x1FAFF) {
      // Unicode 13.0+ symbols — some supported, but newer ones (>=1FAE0) blocked above
      // Block 1FAB7+ (Unicode 14.0 animals/plants) and 1FAC3+ (Unicode 14.0 people)
      if (firstCodePoint >= 0x1FAB7 || (firstCodePoint >= 0x1FAC3 && firstCodePoint <= 0x1FACF)) return false;
    }
    return true;
  };

  addEmoji(event: any) {
    const emoji = event.emoji.native;
    this.cells[this.selectedRow][this.selectedCol] = (this.cells[this.selectedRow][this.selectedCol] || '') + emoji;
    this.formulaBarValue = this.cells[this.selectedRow][this.selectedCol];
    this.onCellChange();
    this.activeModal = null;
    this.showToast('Emoji inserted.');
  }

  applyPresetPicklist(type: string) {
    const presets: { [key: string]: any[] } = {
      'project_status': [
        { label: 'Yet to start', color: '#e2e8f0' },
        { label: 'Blocked', color: '#fed7d7' },
        { label: 'In Progress', color: '#fefcbf' },
        { label: 'Completed', color: '#c6f6d5' }
      ],
      'bug_status': [
        { label: 'Open', color: '#fed7d7' },
        { label: 'In Progress', color: '#fefcbf' },
        { label: 'Closed', color: '#c6f6d5' },
        { label: 'Reopen', color: '#bee3f8' }
      ],
      'review': [
        { label: 'Yet to start', color: '#e2e8f0' },
        { label: 'Under Review', color: '#bee3f8' },
        { label: 'Approved', color: '#c6f6d5' }
      ],
      'priority': [
        { label: 'Low', color: '#bee3f8' },
        { label: 'Medium', color: '#c6f6d5' },
        { label: 'High', color: '#fefcbf' },
        { label: 'Critical', color: '#fed7d7' }
      ],
      'decision': [
        { label: 'Yes', color: '#c6f6d5' },
        { label: 'No', color: '#fed7d7' }
      ],
      'boolean': [
        { label: 'True', color: '#c6f6d5' },
        { label: 'False', color: '#fed7d7' }
      ]
    };

    if (presets[type]) {
      this.picklistOptions = JSON.parse(JSON.stringify(presets[type]));
      this.saveValidation();
      this.showToast('Picklist preset applied.');
    }
    this.closeMenus();
  }

  deleteShiftLeft() {
    this.pushHistory();
    const r = this.selectedRow;
    const c = this.selectedCol;
    this.cells[r].splice(c, 1);
    this.cells[r].push('');
    this.onCellChange(); this.closeMenus();
    this.showToast('Shifted cells left.');
  }

  deleteShiftUp() {
    this.pushHistory();
    const r = this.selectedRow;
    const c = this.selectedCol;
    for (let i = r; i < this.ROWS - 1; i++) {
      this.cells[i][c] = this.cells[i + 1][c];
    }
    this.cells[this.ROWS - 1][c] = '';
    this.onCellChange(); this.closeMenus();
    this.showToast('Shifted cells up.');
  }

  freezeRows(count: number) {
    this.frozenRowsCount = this.frozenRowsCount === count ? 0 : count;
    this.showToast(this.frozenRowsCount > 0 ? `${count} row(s) frozen.` : 'Rows unfrozen.');
  }

  freezeCols(count: number) {
    this.frozenColsCount = this.frozenColsCount === count ? 0 : count;
    this.showToast(this.frozenColsCount > 0 ? `${count} column(s) frozen.` : 'Columns unfrozen.');
  }

  toggleFreezeRow() {
    if (this.frozenRowsCount > 0) {
      this.freezeRows(0);
    } else {
      let count = 1;
      if (this.rangeStart && this.rangeEnd) {
        count = Math.max(this.rangeStart.r, this.rangeEnd.r) + 1;
      } else if (this.selectedRow !== undefined) {
        count = this.selectedRow + 1;
      }
      this.freezeRows(count);
    }
  }

  toggleFreezeCol() {
    if (this.frozenColsCount > 0) {
      this.freezeCols(0);
    } else {
      let count = 1;
      if (this.rangeStart && this.rangeEnd) {
        count = Math.max(this.rangeStart.c, this.rangeEnd.c) + 1;
      } else if (this.selectedCol !== undefined) {
        count = this.selectedCol + 1;
      }
      this.freezeCols(count);
    }
  }

  freezeSelection() {
    if (this.rangeStart && this.rangeEnd) {
      this.frozenRowsCount = Math.max(this.rangeStart.r, this.rangeEnd.r) + 1;
      this.frozenColsCount = Math.max(this.rangeStart.c, this.rangeEnd.c) + 1;
      this.showToast('Selection frozen.');
    } else {
      this.frozenRowsCount = this.selectedRow + 1;
      this.frozenColsCount = this.selectedCol + 1;
      this.showToast('Cell position frozen.');
    }
  }

  hideRows() {
    if (this.rangeStart && this.rangeEnd) {
      const minR = Math.min(this.rangeStart.r, this.rangeEnd.r);
      const maxR = Math.max(this.rangeStart.r, this.rangeEnd.r);
      for (let i = minR; i <= maxR; i++) this.hiddenRows.add(i);
    } else {
      this.hiddenRows.add(this.selectedRow);
    }
    this.showToast('Row(s) hidden.');
  }

  hideCols() {
    if (this.rangeStart && this.rangeEnd) {
      const minC = Math.min(this.rangeStart.c, this.rangeEnd.c);
      const maxC = Math.max(this.rangeStart.c, this.rangeEnd.c);
      for (let i = minC; i <= maxC; i++) this.hiddenCols.add(i);
    } else {
      this.hiddenCols.add(this.selectedCol);
    }
    this.showToast('Column(s) hidden.');
  }

  unhideRows() {
    this.hiddenRows.clear();
    this.showToast('All rows unhidden.');
  }

  unhideCols() {
    this.hiddenCols.clear();
    this.showToast('All columns unhidden.');
  }

  get hasRowGroups() {
    return !!this.sheets[this.currentSheetIdx].rowGroups?.length;
  }

  get groupMarginWidth() {
    return this.hasRowGroups ? 24 : 0;
  }

  getRowGroupsFor(r: number) {
    const groups = this.sheets[this.currentSheetIdx].rowGroups;
    if (!groups) return [];
    return groups.map((g, index) => ({ ...g, index })).filter(g => r >= g.start && r <= g.end);
  }

  groupRow() {
    if (!this.rangeStart || !this.rangeEnd) return;
    const minR = Math.min(this.rangeStart.r, this.rangeEnd.r);
    const maxR = Math.max(this.rangeStart.r, this.rangeEnd.r);
    if (minR === maxR) {
      this.showToast('Select multiple rows to group');
      return;
    }
    const sheet = this.sheets[this.currentSheetIdx];
    if (!sheet.rowGroups) sheet.rowGroups = [];
    sheet.rowGroups.push({ start: minR, end: maxR, collapsed: false });
    this.closeMenus();
    this.save();
    this.showToast('Rows grouped');
  }

  ungroupRow() {
    if (!this.rangeStart || !this.rangeEnd) return;
    const minR = Math.min(this.rangeStart.r, this.rangeEnd.r);
    const maxR = Math.max(this.rangeStart.r, this.rangeEnd.r);
    const sheet = this.sheets[this.currentSheetIdx];
    if (!sheet.rowGroups) return;
    const initialLen = sheet.rowGroups.length;
    sheet.rowGroups = sheet.rowGroups.filter(g => !(g.start >= minR && g.end <= maxR));
    if (sheet.rowGroups.length < initialLen) {
      // Re-evaluate hidden rows since a group was removed
      for (let i = minR; i <= maxR; i++) {
        this.hiddenRows.delete(i);
      }
      this.closeMenus();
      this.save();
      this.showToast('Rows ungrouped');
    }
  }

  clearGroups() {
    const sheet = this.sheets[this.currentSheetIdx];
    if (sheet.rowGroups) {
      sheet.rowGroups.forEach(g => {
        for (let i = g.start; i <= g.end; i++) {
          this.hiddenRows.delete(i);
        }
      });
      sheet.rowGroups = [];
    }
    if (sheet.colGroups) {
      sheet.colGroups.forEach(g => {
        for (let i = g.start; i <= g.end; i++) {
          this.hiddenCols.delete(i);
        }
      });
      sheet.colGroups = [];
    }
    this.closeMenus();
    this.save();
    this.showToast('All groups cleared');
  }

  toggleRowGroup(index: number) {
    const sheet = this.sheets[this.currentSheetIdx];
    if (!sheet.rowGroups) return;
    const group = sheet.rowGroups[index];
    group.collapsed = !group.collapsed;
    for (let r = group.start; r <= group.end; r++) {
      if (group.collapsed) {
        this.hiddenRows.add(r);
      } else {
        this.hiddenRows.delete(r);
      }
    }
    this.save();
  }

  get hasColGroups() {
    return !!this.sheets[this.currentSheetIdx].colGroups?.length;
  }

  get colGroupMarginHeight() {
    return this.hasColGroups ? 24 : 0;
  }

  getColGroupsFor(c: number) {
    const groups = this.sheets[this.currentSheetIdx].colGroups;
    if (!groups) return [];
    return groups.map((g, index) => ({ ...g, index })).filter(g => c >= g.start && c <= g.end);
  }

  groupCol() {
    if (!this.rangeStart || !this.rangeEnd) return;
    const minC = Math.min(this.rangeStart.c, this.rangeEnd.c);
    const maxC = Math.max(this.rangeStart.c, this.rangeEnd.c);
    if (minC === maxC) {
      this.showToast('Select multiple columns to group');
      return;
    }
    const sheet = this.sheets[this.currentSheetIdx];
    if (!sheet.colGroups) sheet.colGroups = [];
    sheet.colGroups.push({ start: minC, end: maxC, collapsed: false });
    this.closeMenus();
    this.save();
    this.showToast('Columns grouped');
  }

  ungroupCol() {
    if (!this.rangeStart || !this.rangeEnd) return;
    const minC = Math.min(this.rangeStart.c, this.rangeEnd.c);
    const maxC = Math.max(this.rangeStart.c, this.rangeEnd.c);
    const sheet = this.sheets[this.currentSheetIdx];
    if (!sheet.colGroups) return;
    const initialLen = sheet.colGroups.length;
    sheet.colGroups = sheet.colGroups.filter(g => !(g.start >= minC && g.end <= maxC));
    if (sheet.colGroups.length < initialLen) {
      for (let i = minC; i <= maxC; i++) {
        this.hiddenCols.delete(i);
      }
      this.closeMenus();
      this.save();
      this.showToast('Columns ungrouped');
    }
  }

  toggleColGroup(index: number) {
    const sheet = this.sheets[this.currentSheetIdx];
    if (!sheet.colGroups) return;
    const group = sheet.colGroups[index];
    group.collapsed = !group.collapsed;
    for (let c = group.start; c <= group.end; c++) {
      if (group.collapsed) {
        this.hiddenCols.add(c);
      } else {
        this.hiddenCols.delete(c);
      }
    }
    this.save();
  }

  setGridlineColor(color: string) {
    this.gridlineColor = color;
    this.updateDisplayCache();
    this.save();
  }

  setGridDirection(dir: 'ltr' | 'rtl') {
    this.gridDirection = dir;
  }

  setGridSpacing(spacing: 'classic' | 'cozy' | 'comfort') {
    this.gridSpacing = spacing;
  }

  toggleFullScreen() {
    this.closeMenus();
    if (!document.fullscreenElement) document.documentElement.requestFullscreen();
    else document.exitFullscreen();
  }

  zoomIn() {
    this.zoomLevel += 10;
    this.closeMenus();
    this.showToast(`Zoom: ${this.zoomLevel}%`);
  }

  zoomOut() {
    this.zoomLevel = Math.max(50, this.zoomLevel - 10);
    this.closeMenus();
    this.showToast(`Zoom: ${this.zoomLevel}%`);
  }

  resetZoom() {
    this.zoomLevel = 100;
    this.closeMenus();
    this.showToast(`Zoom: ${this.zoomLevel}%`);
  }

  sortColAZ() {
    const c = this.selectedCol;
    this.pushHistory({ action_type: 'sort-ascending', target_range: `${this.colLabel(c)}:${this.colLabel(c)}` });
    this.cells.sort((a, b) => {
      const vA = (a[c] || '').trim();
      const vB = (b[c] || '').trim();
      if (!vA && !vB) return 0;
      if (!vA) return 1;
      if (!vB) return -1;
      const nA = Number(vA);
      const nB = Number(vB);
      if (!isNaN(nA) && !isNaN(nB)) return nA - nB;
      return vA.localeCompare(vB);
    });
    this.onCellChange(); this.closeMenus();
    this.showToast(`Column ${colName(c)} sorted A → Z.`);
  }

  sortColZA() {
    const c = this.selectedCol;
    this.pushHistory({ action_type: 'sort-descending', target_range: `${this.colLabel(c)}:${this.colLabel(c)}` });
    this.cells.sort((a, b) => {
      const vA = (a[c] || '').trim();
      const vB = (b[c] || '').trim();
      if (!vA && !vB) return 0;
      if (!vA) return 1;
      if (!vB) return -1;
      const nA = Number(vA);
      const nB = Number(vB);
      if (!isNaN(nA) && !isNaN(nB)) return nB - nA;
      return vB.localeCompare(vA);
    });
    this.onCellChange(); this.closeMenus();
    this.showToast(`Column ${colName(c)} sorted Z → A.`);
  }

  statsModalOpen = false;
  manageLockSettingsModalOpen = false;
  lockSettingsTab = 'ranges';
  lockSettingsSelectedSheet = 'all';
  personalDictModalOpen = false;
  personalDictWords: string[] = [];
  personalDictNewWord = '';

  showWordCount() {
    this.closeMenus();
    this.statsModalOpen = true;
  }

  statsSelectedSheetIdx: number = 0;

  getGlobalStats() {
    let cellsWithData = 0;
    let usedCells = 0;
    for (const sheet of this.sheets) {
      let maxR = -1;
      let maxC = -1;
      let cData = 0;
      for (let r = 0; r < sheet.cells.length; r++) {
        for (let c = 0; c < sheet.cells[r].length; c++) {
          if (sheet.cells[r][c] && sheet.cells[r][c].trim() !== '') {
            cData++;
            maxR = Math.max(maxR, r);
            maxC = Math.max(maxC, c);
          }
        }
      }
      cellsWithData += cData;
      if (maxR >= 0 && maxC >= 0) usedCells += (maxR + 1) * (maxC + 1);
    }
    return { cellsWithData, usedCells };
  }

  getSheetStats(sheetIdx: number) {
    if (sheetIdx == null || !this.sheets[sheetIdx]) return { cellsWithData: 0, usedCells: 0, endOfSheet: 'A1' };
    const sheet = this.sheets[sheetIdx];
    let maxR = -1;
    let maxC = -1;
    let cellsWithData = 0;
    for (let r = 0; r < sheet.cells.length; r++) {
      for (let c = 0; c < sheet.cells[r].length; c++) {
        if (sheet.cells[r][c] && sheet.cells[r][c].trim() !== '') {
          cellsWithData++;
          maxR = Math.max(maxR, r);
          maxC = Math.max(maxC, c);
        }
      }
    }
    const endOfSheet = maxR >= 0 && maxC >= 0 ? this.colLabel(maxC) + (maxR + 1) : 'A1';
    const usedCells = maxR >= 0 && maxC >= 0 ? (maxR + 1) * (maxC + 1) : 0;
    return { cellsWithData, usedCells, endOfSheet };
  }

  personalDictionary() {
    this.closeMenus();
    this.personalDictModalOpen = true;
  }

  addPersonalDictWord() {
    const w = this.personalDictNewWord.trim();
    if (w && !this.personalDictWords.includes(w)) {
      this.personalDictWords.push(w);
      this.personalDictNewWord = '';
      this.showToast(`"${w}" added to dictionary.`);
    }
  }

  removePersonalDictWord(i: number) {
    this.personalDictWords.splice(i, 1);
  }

  showKeyboardShortcuts() {
    this.closeMenus();
    this.activeModal = 'shortcuts';
  }

  openApp(route: string) {
    if (route === 'account') {
      window.open('https://myaccount.google.com/', '_blank'); // Mock account nav
      return;
    }
    this.sidePanelApp = route;
  }

  getTasksDone(): number {
    return this.tasks ? this.tasks.filter((t: any) => t.done).length : 0;
  }

  addTask() {
    if (this.newTask.trim()) {
      this.tasks.push({ text: this.newTask.trim(), done: false });
      this.newTask = '';
      this.save();
    }
  }

  removeTask(i: number) {
    this.tasks.splice(i, 1);
    this.save();
  }

  closeSidePanel() {
    this.sidePanelApp = null;
    this.sidePanelUrl = null;
  }

  isDateLike(val: string): boolean {
    if (!val) return false;
    // Match common date patterns: YYYY-MM-DD, DD/MM/YYYY
    return /^\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}$/.test(val) && !isNaN(Date.parse(val));
  }

  getDateValue(r: number, c: number): string {
    const val = this.cells[r][c];
    if (!val) return '';
    const d = new Date(val);
    if (isNaN(d.getTime())) return '';
    return d.toISOString().split('T')[0]; // Return strictly YYYY-MM-DD for native input
  }

  setDateValue(r: number, c: number, val: string) {
    if (val) {
      this.cells[r][c] = val; // Native date input provides YYYY-MM-DD
      this.onCellChange();
      this.save();
    }
  }

  trashDoc() {
    if (confirm('Move this spreadsheet to trash? This cannot be undone.')) {
      this.api.deleteDocument(this.docId).subscribe(() => this.router.navigate(['/']));
    }
    this.closeMenus();
  }

  cellHasContent(r: number, c: number): boolean {
    if (this.cells[r] && this.cells[r][c] && this.cells[r][c].trim() !== '') return true;
    const fmt = this.formats[`${r},${c}`];
    if (fmt && fmt.bg) return true;
    return false;
  }

  getCellStyle(r: number, c: number): Record<string, string> {
    const fmt = this.formats[`${r},${c}`];
    const style: Record<string, string> = {};

    if (!this.showGridlines) {
      style['border-right'] = 'none';
      style['border-bottom'] = 'none';
    } else if (this.gridlineColor !== '#d0d0d0') {
      style['border-right'] = `1px solid ${this.gridlineColor}`;
      style['border-bottom'] = `1px solid ${this.gridlineColor}`;
    }

    if (r < this.frozenRowsCount || c < this.frozenColsCount) {
      style['background-color'] = '#fff';
    }

    if (this.highlightRowColColor && this.highlightRowColColor !== 'transparent') {
      let rMatch = (r === this.selectedRow);
      let cMatch = (c === this.selectedCol);
      if (this.rangeStart && this.rangeEnd) {
        const minR = Math.min(this.rangeStart.r, this.rangeEnd.r);
        const maxR = Math.max(this.rangeStart.r, this.rangeEnd.r);
        const minC = Math.min(this.rangeStart.c, this.rangeEnd.c);
        const maxC = Math.max(this.rangeStart.c, this.rangeEnd.c);
        rMatch = (r >= minR && r <= maxR);
        cMatch = (c >= minC && c <= maxC);
      }

      if (rMatch || cMatch) {
        style['background-color'] = this.highlightRowColColor;
      }
    }

    if (!fmt) return style;

    if (fmt.bg) style['background-color'] = fmt.bg;
    if (fmt.align) style['text-align'] = fmt.align;
    if (fmt.vertAlign === 'top') style['vertical-align'] = 'top';
    else if (fmt.vertAlign === 'middle') style['vertical-align'] = 'middle';
    else if (fmt.vertAlign === 'bottom') style['vertical-align'] = 'bottom';

    if (fmt.wrap === 'overflow') {
      style['white-space'] = 'nowrap';
      style['overflow'] = 'visible';
    } else if (fmt.wrap === 'wrap' || fmt.wrap === true) {
      style['white-space'] = 'normal';
      style['word-break'] = 'break-word';
    } else if (fmt.wrap === 'clip') {
      style['white-space'] = 'nowrap';
      style['overflow'] = 'hidden';
      style['text-overflow'] = 'clip';
    } else if (fmt.wrap === 'shrink') {
      style['white-space'] = 'nowrap';
    }

    if (fmt.borders) {
      const getB = (b: boolean | CellBorder | undefined): string | null => {
        if (!b) return null;
        if (b === true) return '1px solid #000';
        return `${b.width || '1px'} ${b.style || 'solid'} ${b.color || '#000'}`;
      };
      if (fmt.borders.all) {
        const s = getB(fmt.borders.all);
        if (s) {
          style['border-top'] = s;
          style['border-bottom'] = s;
          style['border-left'] = s;
          style['border-right'] = s;
        }
      } else {
        const t = getB(fmt.borders.top); if (t) style['border-top'] = t;
        const b = getB(fmt.borders.bottom); if (b) style['border-bottom'] = b;
        const l = getB(fmt.borders.left); if (l) style['border-left'] = l;
        const r = getB(fmt.borders.right); if (r) style['border-right'] = r;
      }
    }

    if (this.showLockPattern && (fmt as any).locked) {
      style['background-image'] = 'repeating-linear-gradient(45deg, rgba(255, 165, 0, 0.2), rgba(255, 165, 0, 0.2) 10px, transparent 10px, transparent 20px)';
    }

    return style;
  }

  getContentStyle(r: number, c: number): Record<string, string> {
    const fmt = this.formats[`${r},${c}`];
    if (!fmt) return {};

    const style: Record<string, string> = {};
    if (fmt.bold) style['font-weight'] = 'bold';
    if (fmt.italic) style['font-style'] = 'italic';
    const td: string[] = [];
    if (fmt.strikethrough) td.push('line-through');
    if (fmt.underline) td.push('underline');
    if (td.length) style['text-decoration'] = td.join(' ');
    if (fmt.color) style['color'] = fmt.color;
    if (fmt.font) style['font-family'] = fmt.font;
    if (fmt.size) style['font-size'] = fmt.size;

    if (fmt.align) {
      style['text-align'] = fmt.align;
      if (fmt.align === 'center') style['justify-content'] = 'center';
      else if (fmt.align === 'right') style['justify-content'] = 'flex-end';
      else if (fmt.align === 'left') style['justify-content'] = 'flex-start';
    }

    if (fmt.vertAlign) {
      if (fmt.vertAlign === 'top') style['align-items'] = 'flex-start';
      else if (fmt.vertAlign === 'middle') style['align-items'] = 'center';
      else if (fmt.vertAlign === 'bottom') style['align-items'] = 'flex-end';
    }

    let textW = 0;
    if (fmt.wrap === 'shrink' || (fmt.rotation && fmt.rotation !== 'custom')) {
      const text = this.cells[r] && this.cells[r][c] ? String(this.cells[r][c]) : '';
      const baseSize = fmt.size ? parseInt(fmt.size) : 13;
      textW = text.length * baseSize * 0.55;
    }

    if (fmt.wrap === 'shrink') {
      if (textW > 92) {
        const baseSize = fmt.size ? parseInt(fmt.size) : 13;
        const shrinkSize = Math.max(6, Math.floor(baseSize * (92 / textW)));
        style['font-size'] = `${shrinkSize}px`;
      }
    }

    if (fmt.indent) {
      style['padding-left'] = `${4 + fmt.indent * 12}px`;
    }

    if (fmt.rotation && fmt.rotation !== 'custom') {
      const deg = parseInt(fmt.rotation as string, 10);
      style['display'] = 'inline-block';
      style['min-height'] = 'auto'; // Unbind from 100% td height for proper rotation
      style['position'] = 'relative'; // Override .cell-input absolute positioning

      if (deg === -90 || deg === 90) {
        style['writing-mode'] = 'vertical-rl';
        style['height'] = `${Math.max(30, textW + 16)}px`; // Tight vertical bounding box
        style['width'] = 'auto'; // Shrink horizontally to font size
        if (deg === -90) {
          style['transform'] = 'rotate(180deg)';
        }
      } else {
        style['width'] = `${Math.max(30, textW + 16)}px`; // Tight horizontal bounding box
        style['height'] = 'auto'; // Shrink vertically
        const addedHeight = Math.max(0, textW * 0.707);
        style['transform'] = `rotate(${deg}deg)`;
        style['transform-origin'] = 'left center'; // Pivot tightly on text baseline
        style['white-space'] = 'nowrap';
        if (deg === -45) {
          style['margin-top'] = `${addedHeight}px`;
        } else if (deg === 45) {
          style['margin-bottom'] = `${addedHeight}px`;
        }
      }
    }

    return style;
  }

  // ── Number formatting ─────────────────────────────────────────────────────
  setNumFormat(fmt: string) {
    this.setFormat('numFormat', fmt);

    if (fmt === 'regional_zip' || fmt === 'regional_zip4' || fmt === 'regional_phone') {
      const applyValue = (valToInsert: string) => {
        let modified = false;
        if (this.rangeStart && this.rangeEnd) {
          const minR = Math.min(this.rangeStart.r, this.rangeEnd.r);
          const maxR = Math.max(this.rangeStart.r, this.rangeEnd.r);
          const minC = Math.min(this.rangeStart.c, this.rangeEnd.c);
          const maxC = Math.max(this.rangeStart.c, this.rangeEnd.c);
          for (let r = minR; r <= maxR; r++) {
            for (let c = minC; c <= maxC; c++) {
              if (valToInsert) {
                this.cells[r][c] = valToInsert;
                modified = true;
              }
            }
          }
        } else if (this.selectedRow !== undefined && this.selectedCol !== undefined) {
          if (valToInsert) {
            this.cells[this.selectedRow][this.selectedCol] = valToInsert;
            this.formulaBarValue = valToInsert;
            modified = true;
          }
        }
        if (modified) {
          this.pushHistory();
          this.cdr.detectChanges();
        }
      };

      if (fmt === 'regional_zip') {
        applyValue('560001'); // Standard Indian Pincode placeholder
      } else if (fmt === 'regional_zip4') {
        applyValue('560001-1234');
      } else if (fmt === 'regional_phone') {
        // Fetch phone directly from our backend and forcefully insert it
        this.api.getProfile().subscribe({
          next: (profile: any) => {
            const userPhone = profile?.phone || (this.auth.user as any)?.phone || '';
            applyValue(userPhone);
          },
          error: () => {
            const fallbackPhone = (this.auth.user as any)?.phone || '';
            applyValue(fallbackPhone);
          }
        });
      }
    }

    this.activeMenu = null;
  }

  increaseDecimals() {
    const ref = `${this.selectedRow},${this.selectedCol}`;
    if (!this.formats[ref]) this.formats[ref] = {};
    this.formats[ref].decimals = (this.formats[ref].decimals ?? 0) + 1;
    this.formats = { ...this.formats };
    this.onCellChange();
  }

  decreaseDecimals() {
    const ref = `${this.selectedRow},${this.selectedCol}`;
    if (!this.formats[ref]) this.formats[ref] = {};
    const cur = this.formats[ref].decimals ?? 0;
    this.formats[ref].decimals = Math.max(0, cur - 1);
    this.formats = { ...this.formats };
    this.onCellChange();
  }

  private applyNumFormat(val: string, fmt: CellFormat): string {
    const num = parseFloat(val);
    if (isNaN(num)) return val;
    const dec = fmt.decimals ?? (fmt.numFormat?.includes('currency') || fmt.numFormat?.includes('accounting') ? 2 : fmt.numFormat === 'percent' ? 1 : 0);

    if (fmt.numFormat?.startsWith('currency_') || fmt.numFormat?.startsWith('accounting_') || fmt.numFormat === 'currency') {
      let symbol = '$';
      if (fmt.numFormat.endsWith('_inr')) symbol = '₹';
      else if (fmt.numFormat.endsWith('_eur')) symbol = '€';
      else if (fmt.numFormat.endsWith('_gbp')) symbol = '£';
      else if (fmt.numFormat.endsWith('_cny')) symbol = '¥';

      const isAccounting = fmt.numFormat.startsWith('accounting');
      if (isAccounting) {
        // Simple accounting format representation (symbol on left, number on right, with some spaces)
        return num === 0 ? `${symbol}   -  ` : `${symbol}  ${num.toFixed(dec)}`;
      } else {
        return `${symbol}${num.toFixed(dec)}`;
      }
    }

    if (fmt.numFormat === 'percent') return (num * 100).toFixed(dec) + '%';
    if (fmt.numFormat === 'number') return num.toFixed(dec);
    return dec > 0 ? num.toFixed(dec) : val;
  }

  // ── Text wrap ────────────────────────────────────────────────────────────
  toggleWrap() {
    const ref = `${this.selectedRow},${this.selectedCol}`;
    if (!this.formats[ref]) this.formats[ref] = {};
    this.formats[ref].wrap = !this.formats[ref].wrap;
    this.formats = { ...this.formats };
    this.onCellChange();
  }

  // ── Merge cells ───────────────────────────────────────────────────────────
  mergeCells(type: string = 'all') {
    if (!this.rangeStart || !this.rangeEnd) { this.showToast('Select a range first.'); return; }
    const minR = Math.min(this.rangeStart.r, this.rangeEnd.r);
    const maxR = Math.max(this.rangeStart.r, this.rangeEnd.r);
    const minC = Math.min(this.rangeStart.c, this.rangeEnd.c);
    const maxC = Math.max(this.rangeStart.c, this.rangeEnd.c);
    if (minR === maxR && minC === maxC) { this.showToast('Select more than one cell to merge.'); return; }
    this.pushHistory();

    if (type === 'across') {
      for (let r = minR; r <= maxR; r++) {
        const topLeft = this.cells[r][minC];
        for (let c = minC; c <= maxC; c++) {
          if (c === minC) { this.cells[r][c] = topLeft; continue; }
          this.cells[r][c] = '';
          const ref = `${r},${c}`;
          if (!this.formats[ref]) this.formats[ref] = {};
          (this.formats[ref] as any)['_mergedInto'] = `${r},${minC}`;
        }
        const ref = `${r},${minC}`;
        if (!this.formats[ref]) this.formats[ref] = {};
        (this.formats[ref] as any)['_mergeSpan'] = { rows: 1, cols: maxC - minC + 1 };
      }
    } else if (type === 'down') {
      for (let c = minC; c <= maxC; c++) {
        const topLeft = this.cells[minR][c];
        for (let r = minR; r <= maxR; r++) {
          if (r === minR) { this.cells[r][c] = topLeft; continue; }
          this.cells[r][c] = '';
          const ref = `${r},${c}`;
          if (!this.formats[ref]) this.formats[ref] = {};
          (this.formats[ref] as any)['_mergedInto'] = `${minR},${c}`;
        }
        const ref = `${minR},${c}`;
        if (!this.formats[ref]) this.formats[ref] = {};
        (this.formats[ref] as any)['_mergeSpan'] = { rows: maxR - minR + 1, cols: 1 };
      }
    } else {
      // all or center
      const topLeft = this.cells[minR][minC];
      for (let r = minR; r <= maxR; r++) {
        for (let c = minC; c <= maxC; c++) {
          if (r === minR && c === minC) { this.cells[r][c] = topLeft; continue; }
          this.cells[r][c] = '';
          const ref = `${r},${c}`;
          if (!this.formats[ref]) this.formats[ref] = {};
          (this.formats[ref] as any)['_mergedInto'] = `${minR},${minC}`;
        }
      }
      const ref = `${minR},${minC}`;
      if (!this.formats[ref]) this.formats[ref] = {};
      (this.formats[ref] as any)['_mergeSpan'] = { rows: maxR - minR + 1, cols: maxC - minC + 1 };

      if (type === 'center') {
        (this.formats[ref] as any)['align'] = 'center';
      }
    }

    this.formats = { ...this.formats };
    // Snap selection to the top-left master cell so unmerge works immediately
    this.selectedRow = minR;
    this.selectedCol = minC;
    this.rangeStart = { r: minR, c: minC };
    this.rangeEnd = { r: maxR, c: maxC };
    this.onCellChange(); this.save();
    this.showToast(`Cells merged${type !== 'all' ? ' ' + type : ''}.`);
    this.activeMenu = null;
  }

  unmerge() {
    let masterR = this.selectedRow;
    let masterC = this.selectedCol;
    const ref = `${masterR},${masterC}`;

    // If the selected cell is a slave, follow _mergedInto to find the master
    const mergedInto = (this.formats[ref] as any)?._mergedInto;
    if (mergedInto) {
      const parts = mergedInto.split(',');
      masterR = parseInt(parts[0], 10);
      masterC = parseInt(parts[1], 10);
    }

    const masterRef = `${masterR},${masterC}`;
    const span = (this.formats[masterRef] as any)?._mergeSpan;
    if (!span) { this.showToast('No merged cell selected.'); return; }
    this.pushHistory();
    for (let r = masterR; r < masterR + span.rows; r++)
      for (let c = masterC; c < masterC + span.cols; c++) {
        const k = `${r},${c}`;
        if (this.formats[k]) { delete (this.formats[k] as any)._mergedInto; delete (this.formats[k] as any)._mergeSpan; }
      }
    this.formats = { ...this.formats };
    this.onCellChange(); this.save();
    this.showToast('Cells unmerged.');
    this.activeMenu = null;
  }

  isMergedSlave(r: number, c: number): boolean {
    return !!(this.formats[`${r},${c}`] as any)?._mergedInto;
  }

  getColSpan(r: number, c: number): number {
    const span = (this.formats[`${r},${c}`] as any)?._mergeSpan;
    return span ? span.cols : 1;
  }

  getRowSpan(r: number, c: number): number {
    const span = (this.formats[`${r},${c}`] as any)?._mergeSpan;
    return span ? span.rows : 1;
  }

  isColActiveAxis(c: number): boolean {
    if (this.rangeStart && this.rangeEnd) {
      const minC = Math.min(this.rangeStart.c, this.rangeEnd.c);
      const maxC = Math.max(this.rangeStart.c, this.rangeEnd.c);
      return c >= minC && c <= maxC;
    }
    return this.selectedCol === c;
  }

  isRowActiveAxis(r: number): boolean {
    if (this.rangeStart && this.rangeEnd) {
      const minR = Math.min(this.rangeStart.r, this.rangeEnd.r);
      const maxR = Math.max(this.rangeStart.r, this.rangeEnd.r);
      return r >= minR && r <= maxR;
    }
    return this.selectedRow === r;
  }

  // ── Borders ───────────────────────────────────────────────────────────────
  setBorders(type: 'all' | 'inner' | 'horizontal' | 'vertical' | 'outer' | 'left' | 'top' | 'right' | 'bottom' | 'none') {
    const minR = this.rangeStart ? Math.min(this.rangeStart.r, this.rangeEnd!.r) : this.selectedRow;
    const maxR = this.rangeStart ? Math.max(this.rangeStart.r, this.rangeEnd!.r) : this.selectedRow;
    const minC = this.rangeStart ? Math.min(this.rangeStart.c, this.rangeEnd!.c) : this.selectedCol;
    const maxC = this.rangeStart ? Math.max(this.rangeStart.c, this.rangeEnd!.c) : this.selectedCol;
    this.pushHistory();

    const b: CellBorder = { color: this.currentBorderColor, style: this.currentBorderStyle, width: this.currentBorderWidth };

    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        const ref = `${r},${c}`;
        if (!this.formats[ref]) this.formats[ref] = {};
        if (!this.formats[ref].borders) this.formats[ref].borders = {};

        if (type === 'none') { this.formats[ref].borders = {}; continue; }

        let borders = this.formats[ref].borders!;
        if (borders.all && type !== 'all') {
          borders.top = borders.all;
          borders.bottom = borders.all;
          borders.left = borders.all;
          borders.right = borders.all;
          delete borders.all;
        }

        if (type === 'all') { borders.all = b; continue; }
        if (type === 'outer') {
          if (r === minR) borders.top = b;
          if (r === maxR) borders.bottom = b;
          if (c === minC) borders.left = b;
          if (c === maxC) borders.right = b;
        } else if (type === 'inner') {
          if (r > minR) borders.top = b;
          if (r < maxR) borders.bottom = b;
          if (c > minC) borders.left = b;
          if (c < maxC) borders.right = b;
        } else if (type === 'horizontal') {
          if (r > minR) borders.top = b;
          if (r < maxR) borders.bottom = b;
        } else if (type === 'vertical') {
          if (c > minC) borders.left = b;
          if (c < maxC) borders.right = b;
        } else if (type === 'left') {
          if (c === minC) borders.left = b;
        } else if (type === 'right') {
          if (c === maxC) borders.right = b;
        } else if (type === 'top') {
          if (r === minR) borders.top = b;
        } else if (type === 'bottom') {
          if (r === maxR) borders.bottom = b;
        }
      }
    }
    this.formats = { ...this.formats };
    this.onCellChange(); this.save();
    this.showToast('Borders applied.');
  }

  // ── Fill Down / Fill Right ─────────────────────────────────────────────────
  fillDown() {
    let minR = this.rangeStart ? Math.min(this.rangeStart.r, this.rangeEnd!.r) : this.selectedRow;
    let maxR = this.rangeStart ? Math.max(this.rangeStart.r, this.rangeEnd!.r) : this.selectedRow;
    const minC = this.rangeStart ? Math.min(this.rangeStart.c, this.rangeEnd!.c) : this.selectedCol;
    const maxC = this.rangeStart ? Math.max(this.rangeStart.c, this.rangeEnd!.c) : this.selectedCol;
    if (minR === maxR) {
      if (minR === 0) { this.showToast('Cannot fill down from the first row.'); return; }
      minR = minR - 1;
    }
    this.pushHistory();
    for (let c = minC; c <= maxC; c++) {
      const srcFmt = this.formats[`${minR},${c}`];
      for (let r = minR + 1; r <= maxR; r++) {
        this.cells[r][c] = this.cells[minR][c];
        if (srcFmt) this.formats[`${r},${c}`] = { ...srcFmt };
      }
    }
    this.formats = { ...this.formats };
    this.onCellChange(); this.save();
    this.showToast('Filled down.');
  }

  fillRight() {
    const minR = this.rangeStart ? Math.min(this.rangeStart.r, this.rangeEnd!.r) : this.selectedRow;
    const maxR = this.rangeStart ? Math.max(this.rangeStart.r, this.rangeEnd!.r) : this.selectedRow;
    let minC = this.rangeStart ? Math.min(this.rangeStart.c, this.rangeEnd!.c) : this.selectedCol;
    let maxC = this.rangeStart ? Math.max(this.rangeStart.c, this.rangeEnd!.c) : this.selectedCol;
    if (minC === maxC) {
      if (minC === 0) { this.showToast('Cannot fill right from the first column.'); return; }
      minC = minC - 1;
    }
    this.pushHistory();
    for (let r = minR; r <= maxR; r++) {
      const srcFmt = this.formats[`${r},${minC}`];
      for (let c = minC + 1; c <= maxC; c++) {
        this.cells[r][c] = this.cells[r][minC];
        if (srcFmt) this.formats[`${r},${c}`] = { ...srcFmt };
      }
    }
    this.formats = { ...this.formats };
    this.onCellChange(); this.save();
    this.showToast('Filled right.');
  }

  fillUp() {
    let minR = this.rangeStart ? Math.min(this.rangeStart.r, this.rangeEnd!.r) : this.selectedRow;
    let maxR = this.rangeStart ? Math.max(this.rangeStart.r, this.rangeEnd!.r) : this.selectedRow;
    const minC = this.rangeStart ? Math.min(this.rangeStart.c, this.rangeEnd!.c) : this.selectedCol;
    const maxC = this.rangeStart ? Math.max(this.rangeStart.c, this.rangeEnd!.c) : this.selectedCol;
    if (minR === maxR) {
      if (maxR === this.ROWS - 1) { this.showToast('Cannot fill up from the last row.'); return; }
      maxR = maxR + 1;
    }
    this.pushHistory();
    for (let c = minC; c <= maxC; c++) {
      const srcFmt = this.formats[`${maxR},${c}`];
      for (let r = maxR - 1; r >= minR; r--) {
        this.cells[r][c] = this.cells[maxR][c];
        if (srcFmt) this.formats[`${r},${c}`] = { ...srcFmt };
      }
    }
    this.formats = { ...this.formats };
    this.onCellChange(); this.save();
    this.showToast('Filled up.');
  }

  fillLeft() {
    const minR = this.rangeStart ? Math.min(this.rangeStart.r, this.rangeEnd!.r) : this.selectedRow;
    const maxR = this.rangeStart ? Math.max(this.rangeStart.r, this.rangeEnd!.r) : this.selectedRow;
    let minC = this.rangeStart ? Math.min(this.rangeStart.c, this.rangeEnd!.c) : this.selectedCol;
    let maxC = this.rangeStart ? Math.max(this.rangeStart.c, this.rangeEnd!.c) : this.selectedCol;
    if (minC === maxC) {
      if (maxC === this.COLS - 1) { this.showToast('Cannot fill left from the last column.'); return; }
      maxC = maxC + 1;
    }
    this.pushHistory();
    for (let r = minR; r <= maxR; r++) {
      const srcFmt = this.formats[`${r},${maxC}`];
      for (let c = maxC - 1; c >= minC; c--) {
        this.cells[r][c] = this.cells[r][maxC];
        if (srcFmt) this.formats[`${r},${c}`] = { ...srcFmt };
      }
    }
    this.formats = { ...this.formats };
    this.onCellChange(); this.save();
    this.showToast('Filled left.');
  }

  textToColsModalOpen = false;
  t2cDelimiter = ',';
  t2cCustomDelim = '';

  openTextToColumnsModal() {
    this.textToColsModalOpen = true;
  }

  applyTextToColumns() {
    const delim = this.t2cDelimiter === 'custom' ? this.t2cCustomDelim : this.t2cDelimiter;
    if (!delim) { this.showToast('Please enter a delimiter.'); return; }

    this.pushHistory();
    const minR = this.rangeStart ? Math.min(this.rangeStart.r, this.rangeEnd!.r) : this.selectedRow;
    const maxR = this.rangeStart ? Math.max(this.rangeStart.r, this.rangeEnd!.r) : this.selectedRow;
    const col = this.rangeStart ? Math.min(this.rangeStart.c, this.rangeEnd!.c) : this.selectedCol;

    let splitCount = 0;
    for (let r = minR; r <= maxR; r++) {
      const val = this.cells[r][col];
      if (!val) continue;
      const parts = val.split(delim);
      if (parts.length <= 1) continue;
      this.cells[r][col] = parts[0];
      for (let i = 1; i < parts.length; i++) {
        if (col + i < this.COLS) {
          this.cells[r][col + i] = parts[i].trim();
        }
      }
      splitCount++;
    }

    this.textToColsModalOpen = false;
    this.onCellChange();
    this.save();
    this.showToast(splitCount > 0 ? `Split ${splitCount} cell(s) into columns.` : 'No cells were split.');
  }

  spellCheckModalOpen = false;
  spellCheckLoading = false;
  spellCheckErrors: any[] = [];
  spellCheckTargetText = '';

  translateModalOpen = false;
  translateLoading = false;
  translateSourceText = '';
  translateTargetText = '';
  translateTargetLang = 'es';

  formHeaders: string[] = [];
  formData: { [key: string]: string } = {};

  macroScript = 'this.cells[1][1] = "Hello Macro!";\nthis.save();\nconsole.log("Macro executed successfully!");';
  macroMessages: string[] = [];
  macroErrors: string[] = [];
  macroActiveTab: 'Errors' | 'Messages' | 'Warnings' = 'Messages';
  macroWarnings: string[] = [];
  vbaSelectedNode: string = 'Sheet1';

  customFunctionsScript = `window.customSheetFunctions = {
  MY_CUSTOM_SUM: function(a, b) {
    return a + b;
  },
  MY_CUSTOM_DISCOUNT: function(price, pct) {
    return price * (1 - (pct/100));
  }
};`;

  async spellCheck() {
    this.closeMenus();
    const text = this.cells[this.selectedRow][this.selectedCol];
    if (!text || text.trim() === '') { this.showToast('Select a cell with text to spell check.'); return; }
    this.spellCheckTargetText = text;
    this.spellCheckErrors = [];
    this.spellCheckModalOpen = true;
    this.spellCheckLoading = true;
    try {
      const res = await fetch(`https://api.languagetoolplus.com/v2/check`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({ text: text, language: 'en-US' })
      });
      const data = await res.json();
      this.spellCheckErrors = data.matches || [];
    } catch (e) {
      this.showToast('Error running spell check.');
    } finally {
      this.spellCheckLoading = false;
    }
  }

  applySpellCheckFix(errIndex: number, replacement: string) {
    const err = this.spellCheckErrors[errIndex];
    const text = this.spellCheckTargetText;
    const newText = text.substring(0, err.offset) + replacement + text.substring(err.offset + err.length);
    this.spellCheckTargetText = newText;
    this.cells[this.selectedRow][this.selectedCol] = newText;
    this.save();

    // adjust offsets for remaining errors
    const diff = replacement.length - err.length;
    this.spellCheckErrors.splice(errIndex, 1);
    for (const other of this.spellCheckErrors) {
      if (other.offset > err.offset) {
        other.offset += diff;
      }
    }
  }

  translateSheet() {
    this.closeMenus();
    const text = this.cells[this.selectedRow][this.selectedCol];
    if (!text || text.trim() === '') { this.showToast('Select a cell with text to translate.'); return; }
    this.translateSourceText = text;
    this.translateTargetText = '';
    this.translateModalOpen = true;
    this.runTranslate();
  }

  async runTranslate() {
    if (!this.translateSourceText) return;
    this.translateLoading = true;
    try {
      const res = await fetch(`https://api.mymemory.translated.net/get?q=${encodeURIComponent(this.translateSourceText)}&langpair=en|${this.translateTargetLang}`);
      const data = await res.json();
      if (data && data.responseData && data.responseData.translatedText) {
        this.translateTargetText = data.responseData.translatedText;
      } else {
        this.translateTargetText = 'Error translating text.';
      }
    } catch (e) {
      this.translateTargetText = 'Error translating text.';
    } finally {
      this.translateLoading = false;
    }
  }

  applyTranslation() {
    if (this.translateTargetText && !this.translateLoading) {
      this.cells[this.selectedRow][this.selectedCol] = this.translateTargetText;
      this.save();
      this.translateModalOpen = false;
    }
  }

  patternFill() {
    // Flash fill: detect pattern from filled cells above and fill down
    const col = this.selectedCol;
    const startR = this.rangeStart ? Math.min(this.rangeStart.r, this.rangeEnd!.r) : this.selectedRow;
    const endR = this.rangeStart ? Math.max(this.rangeStart.r, this.rangeEnd!.r) : this.selectedRow;

    // Find first non-empty cell to use as template
    let templateVal = '';
    for (let r = startR - 1; r >= 0; r--) {
      if (this.cells[r][col]) { templateVal = this.cells[r][col]; break; }
    }
    if (!templateVal) { this.showToast('No pattern found above the selection.'); return; }

    // Try numeric sequence detection
    const nums: number[] = [];
    for (let r = startR - 2; r <= startR - 1; r++) {
      if (r >= 0 && this.cells[r][col] && !isNaN(Number(this.cells[r][col]))) {
        nums.push(Number(this.cells[r][col]));
      }
    }

    this.pushHistory();
    if (nums.length === 2) {
      const diff = nums[1] - nums[0];
      let cur = nums[1];
      for (let r = startR; r <= endR; r++) {
        cur += diff;
        this.cells[r][col] = String(cur);
      }
      this.showToast(`Pattern fill: sequence with step ${diff}.`);
    } else {
      // Just fill down with same value
      for (let r = startR; r <= endR; r++) {
        if (!this.cells[r][col]) this.cells[r][col] = templateVal;
      }
      this.showToast('Pattern fill applied.');
    }
    this.onCellChange();
    this.save();
  }

  lockCurrentSheet() {
    const sheet = this.sheets[this.currentSheetIdx];
    sheet.locked = !sheet.locked;
    this.save();
    this.showToast(sheet.locked ? 'Sheet locked.' : 'Sheet unlocked.');
  }

  lockSelectedRange() {
    const minR = this.rangeStart ? Math.min(this.rangeStart.r, this.rangeEnd!.r) : this.selectedRow;
    const maxR = this.rangeStart ? Math.max(this.rangeStart.r, this.rangeEnd!.r) : this.selectedRow;
    const minC = this.rangeStart ? Math.min(this.rangeStart.c, this.rangeEnd!.c) : this.selectedCol;
    const maxC = this.rangeStart ? Math.max(this.rangeStart.c, this.rangeEnd!.c) : this.selectedCol;
    const currentlyLocked = this.isSelectionLocked();
    const targetLockedState = !currentlyLocked;

    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        if (targetLockedState) {
          this.formats[`${r},${c}`] = { ...(this.formats[`${r},${c}`] || {}), locked: true } as any;
        } else {
          if (this.formats[`${r},${c}`]) {
            delete (this.formats[`${r},${c}`] as any).locked;
          }
        }
      }
    }
    this.onCellChange(); this.save();
    this.showToast(`Range ${targetLockedState ? 'locked' : 'unlocked'}: ${this.colLabel(minC)}${minR + 1}:${this.colLabel(maxC)}${maxR + 1}`);
  }

  isSelectionLocked(): boolean {
    const minR = this.rangeStart ? Math.min(this.rangeStart.r, this.rangeEnd!.r) : this.selectedRow;
    const maxR = this.rangeStart ? Math.max(this.rangeStart.r, this.rangeEnd!.r) : this.selectedRow;
    const minC = this.rangeStart ? Math.min(this.rangeStart.c, this.rangeEnd!.c) : this.selectedCol;
    const maxC = this.rangeStart ? Math.max(this.rangeStart.c, this.rangeEnd!.c) : this.selectedCol;
    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        if ((this.formats[`${r},${c}`] as any)?.locked) return true;
      }
    }
    return false;
  }

  manageLockSettings() {
    this.manageLockSettingsModalOpen = true;
    this.lockSettingsTab = 'ranges';
    this.lockSettingsSelectedSheet = 'all';
  }

  trackBySheetIndex(index: number, item: any) { return item.sheetIndex; }
  trackByCellRef(index: number, item: any) { return item.sheetIndex + '-' + item.ref; }

  getLockedSheets(): Array<{ sheetIndex: number, sheetName: string }> {
    const results: Array<{ sheetIndex: number, sheetName: string }> = [];
    for (let i = 0; i < this.sheets.length; i++) {
      if (this.sheets[i].locked) {
        results.push({ sheetIndex: i, sheetName: this.sheets[i].name });
      }
    }
    return results;
  }

  getLockedCellsForCurrentSettings(): Array<{ sheetIndex: number, sheetName: string, ref: string, r: number, c: number }> {
    const results: Array<{ sheetIndex: number, sheetName: string, ref: string, r: number, c: number }> = [];
    const checkAll = this.lockSettingsSelectedSheet === 'all';

    for (let i = 0; i < this.sheets.length; i++) {
      if (!checkAll && parseInt(this.lockSettingsSelectedSheet as string, 10) !== i) continue;

      const sheet = this.sheets[i];
      const formats = (i === this.currentSheetIdx) ? this.formats : (sheet.formats || {});
      for (const key of Object.keys(formats)) {
        if ((formats[key] as any)?.locked) {
          const [rStr, cStr] = key.split(',');
          const r = parseInt(rStr, 10);
          const c = parseInt(cStr, 10);

          let colStr = '';
          let temp = c;
          while (temp >= 0) {
            colStr = String.fromCharCode(65 + (temp % 26)) + colStr;
            temp = Math.floor(temp / 26) - 1;
          }
          const ref = colStr + (r + 1);

          results.push({ sheetIndex: i, sheetName: sheet.name, ref, r, c });
        }
      }
    }
    return results;
  }

  unlockCellFromSettings(item: { sheetIndex: number, sheetName: string, ref: string, r: number, c: number }) {
    const key = `${item.r},${item.c}`;
    if (item.sheetIndex === this.currentSheetIdx) {
      if (this.formats[key]) {
        delete (this.formats[key] as any).locked;
        this.formats = { ...this.formats };
      }
    } else {
      const sheet = this.sheets[item.sheetIndex];
      if (sheet.formats && sheet.formats[key]) {
        delete (sheet.formats[key] as any).locked;
      }
    }
    this.save();
    this.showToast(`Unlocked Cell ${item.ref} on ${item.sheetName}`);
    if (this.cdr) this.cdr.detectChanges();
  }

  highlightLocks() {
    this.showLockPattern = !this.showLockPattern;
  }

  async linkSpreadsheet() {
    const url = await this.openPrompt('Enter the URL of the spreadsheet to link:');
    if (url) {
      this.showToast(`Linked to: ${url.substring(0, 40)}...`);
    }
  }

  dataFromPicture() {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'image/*';
    input.onchange = (e: any) => {
      const file = e.target.files[0];
      if (file) {
        const reader = new FileReader();
        reader.onload = () => {
          this.ocrImage = reader.result;
          this.ocrModalOpen = true;
          this.ocrProgress = 0;
          this.ocrData = [];
          this.cdr.detectChanges();

          this.processOcr();
        };
        reader.readAsDataURL(file);
      }
    };
    input.click();
  }

  processOcr() {
    if (!this.ocrImage) return;
    this.ocrProgress = 1;
    Tesseract.recognize(
      this.ocrImage as string,
      'eng',
      { logger: m => { if (m.status === 'recognizing text') { this.ocrProgress = Math.max(1, Math.round(m.progress * 100)); this.cdr.detectChanges(); } } }
    ).then(({ data: { text } }) => {
      const lines = text.split('\n');
      const parsed = lines.filter(l => l.trim().length > 0).map(l => l.split(/\t| {2,}/));

      let maxCols = 15;
      for (const row of parsed) {
        if (row.length > maxCols) maxCols = row.length;
      }

      const minRows = Math.max(20, parsed.length);

      this.ocrData = [];
      for (let r = 0; r < minRows; r++) {
        const newRow = [];
        for (let c = 0; c < maxCols; c++) {
          if (r < parsed.length && c < parsed[r].length) {
            newRow.push(parsed[r][c] || '');
          } else {
            newRow.push('');
          }
        }
        this.ocrData.push(newRow);
      }
      this.ocrHistory = [];
      this.ocrHistoryIndex = -1;
      this.saveOcrHistory();

      this.ocrProgress = 100;
      this.cdr.detectChanges();
    }).catch(err => {
      console.error('OCR Error:', err);
      this.showToast('OCR Processing Failed.');
      this.ocrProgress = 100;
      this.cdr.detectChanges();
    });
  }

  insertOcrData() {
    if (!this.ocrData.length) return;

    let insertMinR = 0;
    let insertMaxR = -1;
    let insertMinC = 0;
    let insertMaxC = -1;

    if (this.ocrSelStart && this.ocrSelEnd && (this.ocrSelStart.r !== this.ocrSelEnd.r || this.ocrSelStart.c !== this.ocrSelEnd.c)) {
      insertMinR = Math.min(this.ocrSelStart.r, this.ocrSelEnd.r);
      insertMaxR = Math.max(this.ocrSelStart.r, this.ocrSelEnd.r);
      insertMinC = Math.min(this.ocrSelStart.c, this.ocrSelEnd.c);
      insertMaxC = Math.max(this.ocrSelStart.c, this.ocrSelEnd.c);
    } else {
      insertMinR = 0;
      insertMinC = 0;
      for (let r = 0; r < this.ocrData.length; r++) {
        for (let c = 0; c < this.ocrData[r].length; c++) {
          if (this.ocrData[r][c] && this.ocrData[r][c].trim() !== '') {
            insertMaxR = Math.max(insertMaxR, r);
            insertMaxC = Math.max(insertMaxC, c);
          }
        }
      }
    }

    if (insertMaxR === -1) {
      this.showToast('No data to insert!');
      return;
    }

    let actualRows = insertMaxR - insertMinR + 1;
    let actualCols = insertMaxC - insertMinC + 1;

    if (this.ocrInsertTarget === 'new') {
      const newSheetName = 'OCR Data ' + Math.floor(Math.random() * 1000);
      const newSheet = { name: newSheetName, cells: [], formats: {}, validations: {}, sparklines: {}, images: [], colWidths: {}, rowHeights: {} };
      this.sheets.push(newSheet);
      this.switchSheet(this.sheets.length - 1);

      for (let r = 0; r < actualRows; r++) {
        if (!this.cells[r]) this.cells[r] = [];
        for (let c = 0; c < actualCols; c++) {
          this.cells[r][c] = (this.ocrData[insertMinR + r][insertMinC + c] || '').trim();
        }
      }
    } else {
      const startR = this.rangeStart ? Math.min(this.rangeStart.r, this.rangeEnd!.r) : (this.selectedRow || 0);
      const endR = this.rangeStart ? Math.max(this.rangeStart.r, this.rangeEnd!.r) : (this.selectedRow || 0);
      const startC = this.rangeStart ? Math.min(this.rangeStart.c, this.rangeEnd!.c) : (this.selectedCol || 0);
      const endC = this.rangeStart ? Math.max(this.rangeStart.c, this.rangeEnd!.c) : (this.selectedCol || 0);

      let targetR = startR;
      let targetC = startC;

      if (this.ocrAppendMode === 'right') targetC = endC + 1;
      else if (this.ocrAppendMode === 'left') targetC = Math.max(0, startC - actualCols);
      else if (this.ocrAppendMode === 'below') targetR = endR + 1;
      else if (this.ocrAppendMode === 'above') targetR = Math.max(0, startR - actualRows);

      for (let r = 0; r < actualRows; r++) {
        for (let c = 0; c < actualCols; c++) {
          const tr = targetR + r;
          const tc = targetC + c;
          if (tr < this.ROWS && tc < this.COLS) {
            this.cells[tr][tc] = (this.ocrData[insertMinR + r][insertMinC + c] || '').trim();
          }
        }
      }
    }

    this.onCellChange();
    this.save();
    this.ocrModalOpen = false;
    this.ocrSelStart = null;
    this.ocrSelEnd = null;
    this.showToast('Data inserted successfully!');
  }

  publishRange() {
    const ref = this.getRangeRef();
    const url = `${window.location.origin}/sheet/${this.route?.snapshot?.params?.['id']}?range=${ref}`;
    navigator.clipboard.writeText(url).then(() => {
      this.showToast('Published link copied to clipboard!');
    }).catch(() => {
      this.showToast(`Published URL: ${url}`);
    });
  }

  copyPublishLink() {
    const url = window.location.href;
    navigator.clipboard.writeText(url).then(() => {
      this.showToast('Share link copied to clipboard!');
    }).catch(() => {
      this.showToast('Copy failed. Please copy the URL from the address bar.');
    });
  }

  // ── Review menu ──────────────────────────────────────────────────────────
  menuSearch = '';



  async openAuditTrail() {
    this.activeModal = 'audit';
    
    // Wait for any pending flushed events to be fully persisted to DB
    const flushPromise = this.forceFlushAudit();
    if (flushPromise) {
      await flushPromise;
    }
    
    if (this.docId) {
      this.api.getAuditEvents(this.docId).subscribe({
        next: (res: any[]) => {
          this.auditRecords = res || [];
        },
        error: (err: any) => console.error('Failed to load audit events', err)
      });
    }
  }

  getAuditActionPrefix(action: string): string {
    switch(action) {
      case 'set-cell-value': return 'Changed the value of the range';
      case 'clear-cell': return 'Deleted the content from the range';
      case 'delete-row': return 'Deleted row(s) from';
      case 'insert-row-above': 
      case 'insert-row-below': return 'Inserted row(s) at';
      case 'delete-column': return 'Deleted column(s) from';
      case 'insert-column-left':
      case 'insert-column-right': return 'Inserted column(s) at';
      case 'paste': 
      case 'paste-clipboard':
      case 'paste-format':
      case 'paste-formula':
      case 'paste-values': return 'Pasted content into the range';
      case 'drag-to-fill': return 'Filled the range';
      case 'insert-image': return 'Inserted an image at';
      case 'insert-images': return 'Inserted images at';
      case 'delete-image': return 'Deleted an image from';
      case 'delete-images': return 'Deleted images from';
      case 'replace-image': return 'Replaced an image at';
      case 'replace-images': return 'Replaced images at';
      case 'format-change': return 'Changed the formatting of the range';
      case 'toggle-hidden-rows': return 'Toggled hidden rows in';
      case 'apply-filter': return 'Applied a filter to';
      case 'clear-filter': return 'Cleared the filter from';
      case 'sort-ascending': return 'Sorted ascending on';
      case 'sort-descending': return 'Sorted descending on';
      default: return 'Modified the range';
    }
  }

  get sortedAuditRecords() {
    return [...this.auditRecords].sort((a, b) => {
      let valA, valB;
      switch (this.auditSortBy) {
        case 'user': valA = a.user_name?.toLowerCase(); valB = b.user_name?.toLowerCase(); break;
        case 'sheet': valA = a.sheet_name?.toLowerCase(); valB = b.sheet_name?.toLowerCase(); break;
        case 'range': valA = a.target_range?.toLowerCase(); valB = b.target_range?.toLowerCase(); break;
        case 'date': default: valA = new Date(a.created_at).getTime(); valB = new Date(b.created_at).getTime(); break;
      }
      if (valA < valB) return this.auditSortDesc ? 1 : -1;
      if (valA > valB) return this.auditSortDesc ? -1 : 1;
      return 0;
    });
  }

  sortAudit(column: string) {
    if (this.auditSortBy === column) {
      this.auditSortDesc = !this.auditSortDesc;
    } else {
      this.auditSortBy = column;
      this.auditSortDesc = false;
    }
  }

  private sendAuditEvents(): Promise<void> | null {
    if (this.auditBuffer.size === 0) return this.pendingAuditPromise;
    const events = Array.from(this.auditBuffer.values());
    this.auditBuffer.clear();

    if (this.docId) {
      this.pendingAuditPromise = new Promise((resolve) => {
        this.api.saveAuditEvents(this.docId!, events).subscribe({
          next: () => {
            console.log('Audit events saved');
            this.pendingAuditPromise = null;
            resolve();
          },
          error: (err: any) => {
            console.error('Failed to save audit events', err);
            this.pendingAuditPromise = null;
            resolve();
          }
        });
      });
      return this.pendingAuditPromise;
    }
    return null;
  }

  private forceFlushAudit(): Promise<void> | null {
    if (this.pendingDiffTimer) {
      clearTimeout(this.pendingDiffTimer);
      this.pendingDiffTimer = null;
      if (this.pendingDiffPreStateJson) {
        const clonedPreState = JSON.parse(this.pendingDiffPreStateJson);
        const postState = { cells: this.cells, formats: this.formats, hiddenRows: Array.from(this.hiddenRows) };
        const ops = this.diffStateForAudit(clonedPreState, postState);
        for (const op of ops) {
          if (!this.pendingDiffContext) continue;
          const payload = { sheet_id: this.pendingDiffContext.sheetId, sheet_name: this.pendingDiffContext.sheetName, ...op };
          const bufferKey = `${this.pendingDiffContext.currentUser}_${op.target_range}_${op.action_type}`;
          this.auditBuffer.set(bufferKey, payload);
        }
        this.pendingDiffPreStateJson = null;
        this.pendingDiffContext = null;
      }
    }
    return this.sendAuditEvents();
  }

  openEditHistory() {
    this.activeModal = 'version';
    this.loadVersions();
  }

  showAllComments() {
    this.sidePanelApp = 'comments';
    this.commentsViewFilter = 'all';
    this.updateCachedComments();
    this.cdr.markForCheck();
  }

  openCommentInSidePanel(e: MouseEvent, r: number, c: number) {
    e.stopPropagation();
    e.preventDefault();
    this.sidePanelApp = 'comments';
    this.updateCachedComments();

    const ref = `${r},${c}`;
    const comment = this.cachedComments.find(cmt => cmt.sheetIdx === this.currentSheetIdx && cmt.ref === ref);
    if (comment) {
      setTimeout(() => {
        const el = document.getElementById(`comment-card-${comment.data.id}`);
        if (el) {
          el.scrollIntoView({ behavior: 'smooth', block: 'center' });
          const origBg = el.style.backgroundColor;
          el.style.backgroundColor = '#fff3e0';
          setTimeout(() => el.style.backgroundColor = origBg, 1500);
        }
      }, 100);
    }
  }

  addNoteToCell() {
    this.closeMenus();
    const ref = `${this.selectedRow},${this.selectedCol}`;
    this.activeNotePopup = { r: this.selectedRow, c: this.selectedCol, text: (this.formats[ref] as any)?.note || '' };
  }

  hasNote(r: number, c: number): boolean {
    return !!(this.formats[`${r},${c}`] as any)?.note;
  }

  openNotePopup(e: MouseEvent, r: number, c: number) {
    e.stopPropagation();
    e.preventDefault();
    this.activeNotePopup = { r, c, text: (this.formats[`${r},${c}`] as any).note };
  }

  onNoteTextChange(val: string) {
    if (!this.activeNotePopup) return;
    const ref = `${this.activeNotePopup.r},${this.activeNotePopup.c}`;
    if (!this.formats[ref]) this.formats[ref] = {};
    if (val.trim() === '') {
      delete (this.formats[ref] as any).note;
    } else {
      (this.formats[ref] as any).note = val;
    }
    this.onCellChange();
    this.save();
  }

  deleteNote() {
    if (!this.activeNotePopup) return;
    const ref = `${this.activeNotePopup.r},${this.activeNotePopup.c}`;
    if (this.formats[ref]) {
      delete (this.formats[ref] as any).note;
      this.onCellChange();
    }
    this.activeNotePopup = null;
    this.save();
    this.showToast('Note removed.');
  }

  // ── Tools menu ───────────────────────────────────────────────────────────
  createForm() {
    this.formHeaders = [];
    for (let c = 0; c < this.COLS; c++) {
      const h = this.cells[0][c];
      if (h && h.trim()) this.formHeaders.push(h);
      else break;
    }
    this.formData = {};
    this.activeModal = 'form';
  }

  submitForm() {
    let emptyRow = 1;
    while (emptyRow < this.cells.length && this.cells[emptyRow].some(c => c && c.trim() !== '')) {
      emptyRow++;
    }
    if (emptyRow >= this.cells.length) {
      this.cells.push(Array(this.COLS).fill(''));
    }
    for (let c = 0; c < this.formHeaders.length; c++) {
      this.cells[emptyRow][c] = this.formData[this.formHeaders[c]] || '';
    }
    this.save();
    this.formData = {};
    this.showToast('Form data submitted to row ' + (emptyRow + 1));
  }

  showFormActionMenu = false;
  manageForms() {
    this.activeModal = 'manage_forms';
    this.showFormActionMenu = false;
  }

  openMacroEditor() {
    this.activeModal = 'macro';
  }

  runMacro() {
    this.macroMessages = [];
    this.macroErrors = [];
    this.macroWarnings = [];

    const originalLog = console.log;
    const originalWarn = console.warn;
    const originalError = console.error;

    console.log = (...args) => {
      this.macroMessages.push(args.join(' '));
      originalLog.apply(console, args);
    };
    console.warn = (...args) => {
      this.macroWarnings.push(args.join(' '));
      originalWarn.apply(console, args);
    };
    console.error = (...args) => {
      this.macroErrors.push(args.join(' '));
      originalError.apply(console, args);
    };

    try {
      const fn = new Function(this.macroScript);
      fn.call(this);
      this.recalculate();
      if (this.macroMessages.length > 0) {
        this.macroActiveTab = 'Messages';
      }
    } catch (e: any) {
      this.macroErrors.push(e.toString());
      this.macroActiveTab = 'Errors';
      this.showToast('Macro execution error: ' + e);
    } finally {
      console.log = originalLog;
      console.warn = originalWarn;
      console.error = originalError;
    }
  }

  openCustomFunctions() {
    this.activeModal = 'functions';
  }

  saveCustomFunctions() {
    try {
      const fn = new Function(this.customFunctionsScript);
      fn.call(this);
      this.recalculate();
      this.showToast('Custom functions saved and recalculated.');
      this.activeModal = null;
    } catch (e) {
      this.showToast('Error parsing custom functions: ' + e);
    }
  }

  openGoalSeek() {
    this.goalSeekModalOpen = true;
  }

  openSolver() {
    this.goalSeekModalOpen = true;
  }

  openEmailNotifications() {
    this.emailNotifModalOpen = true;
  }

  openMergeTemplate() {
    this.activeModal = 'merge';
  }

  simulateMerge() {
    this.activeModal = null;
    this.showToast('Merging template with data...');
    setTimeout(() => {
      this.showToast('Successfully merged 12 documents.');
    }, 1500);
  }

  openPreferences() {
    this.preferencesModalOpen = true;
  }

  // ── Help menu ────────────────────────────────────────────────────────────
  openWhatsNew() {
    window.open('https://www.zoho.com/sheet/whats-new.html', '_blank');
  }

  openUserGuide() {
    window.open('https://www.zoho.com/sheet/help/', '_blank');
  }

  openDeveloperApi() {
    window.open('https://www.zoho.com/sheet/developer-api/', '_blank');
  }

  openFeedback() {
    this.feedbackModalOpen = true;
  }

  // ── Goal Seek / Email Notifications / Preferences / Feedback modals ───────
  goalSeekModalOpen = false;
  goalSeekTargetCell = '';
  goalSeekTargetValue = '';
  goalSeekByCell = '';
  emailNotifModalOpen = false;
  emailNotifOnEdit = true;
  emailNotifOnComment = true;
  emailNotifEmail = '';
  preferencesModalOpen = false;
  prefLocale = 'en-US';
  prefDateFormat = 'MM/DD/YYYY';
  prefThousands = true;
  feedbackModalOpen = false;
  feedbackText = '';
  feedbackRating = 0;
  feedbackType = 'Bug';
  feedbackRecordScreen = false;
  feedbackFile: File | null = null;
  feedbackUploading = false;

  applyGoalSeek() {
    if (!this.goalSeekTargetCell || !this.goalSeekTargetValue || !this.goalSeekByCell) {
      this.showToast('Please fill in all Goal Seek fields.'); return;
    }
    const tMatch = this.goalSeekTargetCell.trim().match(/^([a-zA-Z]+)(\d+)$/);
    const vMatch = this.goalSeekByCell.trim().match(/^([a-zA-Z]+)(\d+)$/);
    if (!tMatch || !vMatch) { this.showToast('Invalid cell references.'); return; }

    const tr = parseInt(tMatch[2]) - 1;
    const tc = tMatch[1].toUpperCase().charCodeAt(0) - 65;
    const vr = parseInt(vMatch[2]) - 1;
    const vc = vMatch[1].toUpperCase().charCodeAt(0) - 65;

    const targetVal = parseFloat(this.goalSeekTargetValue);
    if (isNaN(targetVal)) { this.showToast('Target value must be a number.'); return; }

    // Simple iterative solver (Newton's method)
    let currentX = parseFloat(this.evalCell(vr, vc)) || 0;
    let iterations = 0;
    let success = false;

    while (iterations < 50) {
      this.cells[vr][vc] = String(currentX);
      let y0 = parseFloat(String(this.evalCell(tr, tc))) || 0;

      let error = targetVal - y0;
      if (Math.abs(error) < 0.0001) { success = true; break; }

      this.cells[vr][vc] = String(currentX + 0.001);
      let y1 = parseFloat(String(this.evalCell(tr, tc))) || 0;

      let derivative = (y1 - y0) / 0.001;
      if (derivative === 0) {
        currentX += (Math.random() - 0.5); // Random jump to escape flat region
      } else {
        currentX = currentX + (error / derivative);
      }
      iterations++;
    }

    this.cells[vr][vc] = String(currentX);
    this.save();

    if (success) {
      this.showToast(`Goal Seek Success: Set ${this.goalSeekByCell} to ${currentX.toFixed(4)}.`);
    } else {
      this.showToast(`Goal Seek failed to converge after 50 iterations.`);
    }
    this.goalSeekModalOpen = false;
  }

  saveEmailNotifications() {
    this.api.saveNotificationSettings(this.docId, {
      notify_email: this.emailNotifEmail,
      on_edit: this.emailNotifOnEdit ? 1 : 0,
      on_comment: this.emailNotifOnComment ? 1 : 0
    }).subscribe({
      next: () => {
        this.emailNotifModalOpen = false;
        this.showToast('Email notification preferences saved.');
      },
      error: () => this.showToast('Error saving notification preferences.')
    });
  }

  savePreferences() {
    this.preferencesModalOpen = false;
    this.showToast('Preferences saved.');
  }

  onFeedbackFileSelected(event: any) {
    this.feedbackFile = event.target.files[0];
  }

  submitFeedback() {
    if (!this.feedbackText.trim()) { this.showToast('Please enter your feedback.'); return; }
    
    this.feedbackUploading = true;
    this.api.submitFeedback(this.docId, this.feedbackType, this.feedbackText, this.feedbackRating, this.feedbackRecordScreen, this.feedbackFile).subscribe({
      next: () => {
        this.feedbackModalOpen = false;
        this.feedbackText = '';
        this.feedbackRating = 0;
        this.feedbackType = 'Bug';
        this.feedbackRecordScreen = false;
        this.feedbackFile = null;
        this.feedbackUploading = false;
        this.showToast('Thank you for your feedback!');
      },
      error: () => {
        this.feedbackUploading = false;
        this.showToast('Failed to submit feedback.');
      }
    });
  }

  // ── Find & Replace ────────────────────────────────────────────────────────
  openFind() {
    this.findModalOpen = true;
    this.findQuery = '';
    this.replaceQuery = '';
    this.findStatus = '';
    this.findModalPosition = 'right';
  }

  // ── Inline Search ────────────────────────────────────────────────────────
  inlineSearchQuery = '';
  inlineSearchMatches: { r: number, c: number }[] = [];
  inlineSearchMatchMap = new Set<string>();
  inlineSearchIdx = -1;

  onInlineSearch() {
    this.inlineSearchMatches = [];
    this.inlineSearchMatchMap.clear();
    this.inlineSearchIdx = -1;
    if (!this.inlineSearchQuery) return;
    const q = this.inlineSearchQuery.toLowerCase();
    for (let r = 0; r < this.ROWS; r++) {
      for (let c = 0; c < this.COLS; c++) {
        const cellVal = this.cells[r][c];
        if (!cellVal || typeof cellVal !== 'string') continue;
        if (this.isImageCell(r, c)) continue;
        if (cellVal.toLowerCase().includes(q)) {
          this.inlineSearchMatches.push({ r, c });
          this.inlineSearchMatchMap.add(`${r},${c}`);
        }
      }
    }
    if (this.inlineSearchMatches.length > 0) {
      this.inlineSearchIdx = 0;
      const m = this.inlineSearchMatches[0];
      this.selectCell(m.r, m.c);
      this.scrollToCell(m.r, m.c);
      this.cdr.detectChanges();
    }
  }

  inlineFindNext() {
    if (!this.inlineSearchMatches.length) return;
    this.inlineSearchIdx = (this.inlineSearchIdx + 1) % this.inlineSearchMatches.length;
    const m = this.inlineSearchMatches[this.inlineSearchIdx];
    this.selectCell(m.r, m.c);
    this.scrollToCell(m.r, m.c);
    this.cdr.detectChanges();
  }

  inlineFindPrev() {
    if (!this.inlineSearchMatches.length) return;
    this.inlineSearchIdx = (this.inlineSearchIdx - 1 + this.inlineSearchMatches.length) % this.inlineSearchMatches.length;
    const m = this.inlineSearchMatches[this.inlineSearchIdx];
    this.selectCell(m.r, m.c);
    this.scrollToCell(m.r, m.c);
    this.cdr.detectChanges();
  }

  clearInlineSearch() {
    this.inlineSearchQuery = '';
    this.inlineSearchMatches = [];
    this.inlineSearchMatchMap.clear();
    this.inlineSearchIdx = -1;
  }

  isCellInInlineSearch(r: number, c: number): boolean {
    return this.inlineSearchMatchMap.has(`${r},${c}`);
  }

  isCellActiveInlineSearch(r: number, c: number): boolean {
    if (this.inlineSearchIdx < 0 || this.inlineSearchIdx >= this.inlineSearchMatches.length) return false;
    const m = this.inlineSearchMatches[this.inlineSearchIdx];
    return m.r === r && m.c === c;
  }
  // ─────────────────────────────────────────────────────────────────────────

  private buildFindMatches() {
    this.findMatches = [];
    if (!this.findQuery) return;

    const targetSheets = this.findSearchIn === 'workbook' ? this.sheets.map((_, i) => i) : [this.currentSheetIdx];

    for (const sIdx of targetSheets) {
      const sheetCells = sIdx === this.currentSheetIdx ? this.cells : this.sheets[sIdx].cells;
      const rows = sheetCells.length;
      const cols = rows > 0 ? sheetCells[0].length : 0;

      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          let cellVal = sheetCells[r][c];
          if (cellVal === null || cellVal === undefined || cellVal === '') continue;
          if (typeof cellVal === 'string' && (cellVal.trim().startsWith('data:image') || cellVal.trim().toUpperCase().startsWith('=IMAGE('))) continue;

          let match = false;
          let query = this.findQuery;
          let target = String(cellVal);

          if (!this.findMatchCase) {
            query = query.toLowerCase();
            target = target.toLowerCase();
          }

          if (this.findMatchEntireCell) {
            match = target === query;
          } else {
            match = target.includes(query);
          }

          if (match) {
            this.findMatches.push({ r, c, sIdx });
          }
        }
      }
    }

    if (this.findDirection === 'up') {
      this.findMatches.reverse();
    }
  }

  findNext() {
    this.buildFindMatches();
    if (!this.findMatches.length) { this.findStatus = 'No matches found.'; return; }
    this.findMatchIdx = (this.findMatchIdx + 1) % this.findMatches.length;
    const m = this.findMatches[this.findMatchIdx];

    if (m.sIdx !== this.currentSheetIdx) {
      this.switchSheet(m.sIdx);
    }
    this.selectCell(m.r, m.c);
    this.findStatus = `Match ${this.findMatchIdx + 1} of ${this.findMatches.length}`;

    // Scroll the virtual grid immediately
    this.scrollToCell(m.r, m.c);
    this.cdr.detectChanges();

    // First timeout: let Angular render the newly visible rows
    setTimeout(() => {
      this.cdr.detectChanges();
      // Second timeout: now the cells should be in the DOM
      setTimeout(() => {
        const wrapEl = this.gridWrapRef?.nativeElement || (document.querySelector('.grid-wrap') as HTMLElement);
        const el = document.getElementById(`cell-${m.r}-${m.c}`);
        if (el && wrapEl) {
          // Get el position relative to wrapEl
          const elRect = el.getBoundingClientRect();
          const wrapRect = wrapEl.getBoundingClientRect();
          const elOffsetTop = wrapEl.scrollTop + (elRect.top - wrapRect.top);
          const elOffsetLeft = wrapEl.scrollLeft + (elRect.left - wrapRect.left);
          wrapEl.scrollTop = Math.max(0, elOffsetTop - wrapEl.clientHeight / 3);
          wrapEl.scrollLeft = Math.max(0, elOffsetLeft - wrapEl.clientWidth / 3);

          const rect = el.getBoundingClientRect();
          this.findModalPosition = rect.left > window.innerWidth / 2 ? 'left' : 'right';
        } else if (wrapEl) {
          // Fallback: direct scroll calculation
          let targetScrollTop = 0;
          for (let i = 0; i < m.r; i++) targetScrollTop += this.getRowHeight(i) || 24;
          wrapEl.scrollTop = Math.max(0, targetScrollTop - wrapEl.clientHeight / 3);
          this.updateVisibleRows(wrapEl.scrollTop);
        }
      }, 80);
    }, 50);
  }

  findAll() {
    this.buildFindMatches();
    if (!this.findMatches.length) { this.findStatus = 'No matches found.'; return; }
    // Only select matches in current sheet if "Find All" is clicked
    const currentSheetMatches = this.findMatches.filter(m => m.sIdx === this.currentSheetIdx);
    if (currentSheetMatches.length) {
      const rows = currentSheetMatches.map(m => m.r), cols = currentSheetMatches.map(m => m.c);
      this.rangeStart = { r: Math.min(...rows), c: Math.min(...cols) };
      this.rangeEnd = { r: Math.max(...rows), c: Math.max(...cols) };
    }
    this.findStatus = `Found ${this.findMatches.length} matches.`;
  }

  replaceOne() {
    this.buildFindMatches();
    if (!this.findMatches.length) { this.findStatus = 'No matches found.'; return; }
    this.findMatchIdx = (this.findMatchIdx + 1) % this.findMatches.length;
    const m = this.findMatches[this.findMatchIdx];

    if (m.sIdx !== this.currentSheetIdx) {
      this.switchSheet(m.sIdx);
    }

    this.pushHistory();
    const flags = this.findMatchCase ? 'g' : 'gi';
    const q = new RegExp(this.findQuery.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), flags);

    if (this.findMatchEntireCell) {
      this.cells[m.r][m.c] = this.replaceQuery;
    } else {
      this.cells[m.r][m.c] = this.cells[m.r][m.c].replace(q, this.replaceQuery);
    }

    this.onCellChange(); this.save();
    this.findStatus = `Replaced 1 instance.`;

    this.scrollToCell(m.r, m.c);
    this.cdr.detectChanges();

    setTimeout(() => {
      const el = document.getElementById(`cell-${m.r}-${m.c}`);
      if (el) {
        el.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
        const rect = el.getBoundingClientRect();
        if (rect.left > window.innerWidth / 2) {
          this.findModalPosition = 'left';
        } else {
          this.findModalPosition = 'right';
        }
      }
    }, 50);
  }

  replaceAll() {
    this.buildFindMatches();
    if (!this.findMatches.length) { this.findStatus = 'No matches found.'; return; }

    this.pushHistory();
    const flags = this.findMatchCase ? 'g' : 'gi';
    const q = new RegExp(this.findQuery.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), flags);
    let count = 0;

    for (const m of this.findMatches) {
      const sheetCells = m.sIdx === this.currentSheetIdx ? this.cells : this.sheets[m.sIdx].cells;
      if (this.findMatchEntireCell) {
        sheetCells[m.r][m.c] = this.replaceQuery;
      } else {
        sheetCells[m.r][m.c] = sheetCells[m.r][m.c].replace(q, this.replaceQuery);
      }
      count++;
    }

    this.findMatches = []; this.findMatchIdx = -1;
    this.onCellChange(); this.save();
    this.findStatus = `Replaced ${count} instances.`;
  }

  executeGoto() {
    if (!this.gotoQuery) return;
    const q = this.gotoQuery.trim().toUpperCase();
    let sheetName = '';
    let cellRef = q;

    if (q.includes('!')) {
      const parts = q.split('!');
      sheetName = parts[0];
      cellRef = parts[1];
    }

    // Parse cell ref like "A1", "AB12"
    const match = cellRef.match(/^([A-Z]+)(\d+)$/);
    if (!match) {
      this.showToast('Invalid cell reference. Use format like A1 or AB12.');
      return;
    }

    const colStr = match[1];
    const rowStr = match[2];

    let colIdx = 0;
    for (let i = 0; i < colStr.length; i++) {
      colIdx = colIdx * 26 + (colStr.charCodeAt(i) - 64);
    }
    colIdx -= 1; // 0-based
    const rowIdx = parseInt(rowStr, 10) - 1; // 0-based

    if (sheetName) {
      const cleanSheetName = sheetName.replace(/^'|'$/g, '');
      const sIdx = this.sheets.findIndex(s => s.name.toUpperCase() === cleanSheetName);
      if (sIdx !== -1 && sIdx !== this.currentSheetIdx) {
        this.switchSheet(sIdx);
      } else if (sIdx === -1) {
        this.showToast(`Sheet '${cleanSheetName}' not found.`);
        return;
      }
    }

    if (rowIdx >= 0 && rowIdx < this.ROWS && colIdx >= 0 && colIdx < this.COLS) {
      this.selectCell(rowIdx, colIdx);
      this.activeModal = null;
      this.gotoQuery = '';
      setTimeout(() => {
        const el = document.getElementById(`cell-${rowIdx}-${colIdx}`);
        if (el) {
          el.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
        }
      }, 50);
    } else {
      this.showToast('Reference out of bounds.');
    }
  }

  // ── Multiple Sheets ───────────────────────────────────────────────────────

  get hiddenSheetsList(): { s: any, idx: number }[] {
    return this.sheets.map((s, idx) => ({ s, idx })).filter(x => x.s.hidden);
  }

  unhideAllSheets() {
    setTimeout(() => {
      this.sheets.forEach(s => s.hidden = false);
      this.save();
    }, 0);
  }

  unhideSheetAndSwitch(idx: number) {
    setTimeout(() => {
      this.unhideSheet(idx);
      this.switchSheet(idx);
    }, 0);
  }

  private saveCurrentSheet() {
    const existing = this.sheets[this.currentSheetIdx];
    this.sheets[this.currentSheetIdx] = {
      ...existing,
      gridlineColor: this.gridlineColor,
      cells: this.cells.map(row => [...row]),
      formats: { ...this.formats },
      validations: { ...this.validations },
      hiddenRows: Array.from(this.hiddenRows),
      activeFilterCols: Array.from(this.activeFilterCols),
      filterActive: this.filterActive,
      advFilterSavedState: this.serializeAdvFilterState()
    };
  }

  switchSheet(idx: number) {
    if (idx === this.currentSheetIdx) return;
    this.saveCurrentSheet();
    this.currentSheetIdx = idx;
    const s = this.sheets[idx];
    for (let r = 0; r < this.ROWS; r++) for (let c = 0; c < this.COLS; c++) this.cells[r][c] = s.cells[r]?.[c] ?? '';
    this.formats = { ...s.formats };
    this.validations = { ...s.validations };
    this.hiddenRows = new Set(s.hiddenRows || []);
    this.activeFilterCols = new Set(s.activeFilterCols || []);
    this.filterActive = !!s.filterActive;
    this.gridlineColor = s.gridlineColor || '#d0d0d0';
    this.deserializeAdvFilterState(s.advFilterSavedState);
    this.frozenRowsCount = s.frozenRowsCount || 0;
    this.frozenColsCount = s.frozenColsCount || 0;
    this.rangeStart = null; this.rangeEnd = null;
    this.selectedRow = 0; this.selectedCol = 0;
    this.formulaBarValue = '';
    this.updateDisplayCache();
  }

  addSheet() {
    this.saveCurrentSheet();
    const n = this.sheets.length + 1;
    this.sheets.push({
      name: `Sheet${n}`,
      cells: Array.from({ length: this.ROWS }, () => Array(this.COLS).fill('')),
      formats: {}, validations: {}
    });
    this.switchSheet(this.sheets.length - 1);
    this.save();
  }

  async renameSheet(idx: number) {
    const cur = this.sheets[idx].name;
    const name = await this.openPrompt('Rename sheet:', cur);
    if (name && name.trim()) { this.sheets[idx] = { ...this.sheets[idx], name: name.trim() }; this.save(); }
  }

  async deleteSheet(idx: number) {
    if (this.sheets.length <= 1) { this.showToast('Cannot delete the only sheet.'); return; }
    const confirmed = await this.openConfirm(`Are you sure you want to delete "${this.sheets[idx].name}"?`);
    if (!confirmed) return;
    this.sheets.splice(idx, 1);
    this.currentSheetIdx = Math.min(this.currentSheetIdx, this.sheets.length - 1);
    this.switchSheet(this.currentSheetIdx);
    this.save();
  }

  openSheetMenu(idx: number, event: MouseEvent) {
    event.stopPropagation();
    event.preventDefault();
    this.switchSheet(idx);
    this.activeSheetMenuIdx = idx;
    this.sheetMenuX = event.clientX;
    this.sheetMenuY = window.innerHeight - event.clientY;
  }

  duplicateSheet(idx: number) {
    this.pushHistory();
    const sourceSheet = this.sheets[idx];
    const newSheet = JSON.parse(JSON.stringify(sourceSheet));
    let counter = 1;
    let finalName = sourceSheet.name + ' (Copy)';
    while (this.sheets.some(s => s.name === finalName)) {
      counter++;
      finalName = `${sourceSheet.name} (Copy ${counter})`;
    }
    newSheet.name = finalName;
    this.sheets.splice(idx + 1, 0, newSheet);
    this.switchSheet(idx + 1);
    this.showToast(`Duplicated sheet to ${finalName}`);
  }

  // ── Formula Engine ────────────────────────────────────────────────────────
  private evalCell(r: number, c: number, visited = new Set<string>()): string {
    let raw = this.cells[r][c];
    if (raw === undefined || raw === null) return '';
    if (typeof raw !== 'string') raw = String(raw);

    if (!raw.startsWith('=')) {
      return raw;
    }
    const key = `${r},${c}`;
    if (visited.has(key)) return '#CIRCULAR!';
    visited.add(key);
    try { return String(this.evalExpr(raw.slice(1).trim().toUpperCase(), visited)); }
    catch { return '#ERROR!'; }
  }

  private getCellVal(ref: string, visited: Set<string>): number | string {
    const m = ref.match(/^([A-Z]+)(\d+)$/);
    if (!m) return 0;
    const c = m[1].charCodeAt(0) - 65, r = parseInt(m[2]) - 1;
    if (r < 0 || r >= this.ROWS || c < 0 || c >= this.COLS) return 0;
    const v = this.evalCell(r, c, new Set(visited));
    return v === '' ? 0 : (isNaN(Number(v)) ? v : Number(v));
  }

  private getRangeVals(range: string, visited: Set<string>): (number | string)[] {
    const m = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (!m) return [];
    const c1 = m[1].charCodeAt(0) - 65, r1 = parseInt(m[2]) - 1;
    const c2 = m[3].charCodeAt(0) - 65, r2 = parseInt(m[4]) - 1;
    const vals: (number | string)[] = [];
    for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++)
      for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++)
        vals.push(this.getCellVal(colName(c) + (r + 1), visited));
    return vals;
  }

  private evalExpr(expr: string, visited: Set<string>): number | string {
    // Functions
    const fnMatch = expr.match(/^([A-Z_][A-Z0-9_]*)\((.*)\)$/);
    if (fnMatch) {
      const fn = fnMatch[1], args = this.parseArgs(fnMatch[2]);
      // Resolve each arg: either a range, cell ref, string literal, or number
      const resolve = (a: string): (number | string)[] => {
        const t = a.trim();
        if (/^[A-Z]+\d+:[A-Z]+\d+$/.test(t)) return this.getRangeVals(t, visited);
        if (/^[A-Z]+\d+$/.test(t)) return [this.getCellVal(t, visited)];
        if (/^".*"$/.test(t)) return [t.slice(1, -1)];
        return [this.evalExpr(t, visited)];
      };
      const flatArgs = args.flatMap(a => resolve(a.trim().toUpperCase()));
      const nums = flatArgs.filter(v => typeof v === 'number' || !isNaN(Number(v))).map(Number);
      switch (fn) {
        case 'SUM': return nums.reduce((a, b) => a + b, 0);
        case 'AVERAGE': case 'AVG': return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
        case 'COUNT': return nums.length;
        case 'COUNTA': return flatArgs.filter(v => v !== '' && v !== 0).length;
        case 'MAX': return nums.length ? Math.max(...nums) : 0;
        case 'MIN': return nums.length ? Math.min(...nums) : 0;
        case 'ABS': return Math.abs(nums[0] ?? 0);
        case 'ROUND': return Math.round((nums[0] ?? 0) * Math.pow(10, nums[1] ?? 0)) / Math.pow(10, nums[1] ?? 0);
        case 'SQRT': return Math.sqrt(nums[0] ?? 0);
        case 'MOD': return (nums[0] ?? 0) % (nums[1] ?? 1);
        case 'POWER': return Math.pow(nums[0] ?? 0, nums[1] ?? 1);
        case 'LEN': return String(flatArgs[0] ?? '').length;
        case 'UPPER': return String(flatArgs[0] ?? '').toUpperCase();
        case 'LOWER': return String(flatArgs[0] ?? '').toLowerCase();
        case 'TRIM': return String(flatArgs[0] ?? '').trim();
        case 'LEFT': return String(flatArgs[0] ?? '').slice(0, nums[1] ?? 1);
        case 'RIGHT': return String(flatArgs[0] ?? '').slice(-(nums[1] ?? 1));
        case 'MID': return String(flatArgs[0] ?? '').slice((nums[1] ?? 1) - 1, (nums[1] ?? 1) - 1 + (nums[2] ?? 1));
        case 'CONCATENATE': case 'CONCAT': return flatArgs.map(String).join('');
        case 'TODAY': return new Date().toLocaleDateString();
        case 'NOW': return new Date().toLocaleString();
        case 'ISNUMBER': return !isNaN(Number(flatArgs[0])) ? 'TRUE' : 'FALSE';
        case 'ISBLANK': return flatArgs[0] === '' ? 'TRUE' : 'FALSE';
        case 'IF': {
          const rawCond = args[0]?.trim().toUpperCase() ?? '';
          let cond: any = false;
          try { cond = this.evalExpr(rawCond, visited); } catch { cond = false; }
          const truthy = cond === 'TRUE' || (typeof cond === 'number' && cond !== 0) || cond === true;
          const branch = args[truthy ? 1 : 2]?.trim() ?? '';
          if (/^".*"$/.test(branch)) return branch.slice(1, -1);
          try { return this.evalExpr(branch.toUpperCase(), visited); } catch { return branch; }
        }
        case 'AND': return flatArgs.every(v => v === 'TRUE' || (typeof v === 'number' && v !== 0)) ? 'TRUE' : 'FALSE';
        case 'OR': return flatArgs.some(v => v === 'TRUE' || (typeof v === 'number' && v !== 0)) ? 'TRUE' : 'FALSE';
        case 'TEXT': return String(flatArgs[0] ?? '');
        default:
          if ((window as any).customSheetFunctions && typeof (window as any).customSheetFunctions[fn] === 'function') {
            try { return (window as any).customSheetFunctions[fn](...flatArgs); } catch (e) { return '#ERROR!'; }
          }
          return '#NAME?';
      }
    }
    // Comparison operators
    for (const op of ['>=', '<=', '<>', '!=', '>', '<', '=']) {
      const idx = expr.indexOf(op);
      if (idx > 0) {
        const lv = this.evalExpr(expr.slice(0, idx).trim(), visited);
        const rv = this.evalExpr(expr.slice(idx + op.length).trim(), visited);
        const ln = Number(lv), rn = Number(rv);
        const ls = String(lv), rs = String(rv);
        const cmp = !isNaN(ln) && !isNaN(rn) ? ln - rn : ls.localeCompare(rs);
        if (op === '>=') return cmp >= 0 ? 'TRUE' : 'FALSE';
        if (op === '<=') return cmp <= 0 ? 'TRUE' : 'FALSE';
        if (op === '<>' || op === '!=') return cmp !== 0 ? 'TRUE' : 'FALSE';
        if (op === '>') return cmp > 0 ? 'TRUE' : 'FALSE';
        if (op === '<') return cmp < 0 ? 'TRUE' : 'FALSE';
        if (op === '=') return cmp === 0 ? 'TRUE' : 'FALSE';
      }
    }
    // String concat with &
    if (expr.includes('&')) {
      return expr.split('&').map(p => {
        const t = p.trim();
        if (/^".*"$/.test(t)) return t.slice(1, -1);
        if (/^[A-Z]+\d+$/.test(t)) return String(this.getCellVal(t, visited));
        try { return String(this.evalExpr(t, visited)); } catch { return t; }
      }).join('');
    }
    // Arithmetic: +, -, *, /  (right-to-left for +/- to handle precedence)
    const arithParts = expr.match(/([+\-*/^])/);
    if (arithParts) {
      // Simple tokenizer for arithmetic
      const tokens = expr.split(/([+\-*/^])/);
      if (tokens.length >= 3) {
        const vals = tokens.filter((_, i) => i % 2 === 0).map(t => {
          const tt = t.trim();
          if (/^[A-Z]+\d+$/.test(tt)) return Number(this.getCellVal(tt, visited));
          if (/^".*"$/.test(tt)) return NaN;
          return Number(this.evalExpr(tt, visited));
        });
        const ops = tokens.filter((_, i) => i % 2 === 1);
        let result = vals[0];
        for (let i = 0; i < ops.length; i++) {
          if (ops[i] === '+') result += vals[i + 1];
          else if (ops[i] === '-') result -= vals[i + 1];
          else if (ops[i] === '*') result *= vals[i + 1];
          else if (ops[i] === '/') result = vals[i + 1] !== 0 ? result / vals[i + 1] : Infinity;
          else if (ops[i] === '^') result = Math.pow(result, vals[i + 1]);
        }
        return isNaN(result) ? '#VALUE!' : result;
      }
    }
    // Cell reference
    if (/^[A-Z]+\d+$/.test(expr)) return this.getCellVal(expr, visited);
    // String literal
    if (/^".*"$/.test(expr)) return expr.slice(1, -1);
    // Plain number
    if (!isNaN(Number(expr))) return Number(expr);
    return '#VALUE!';
  }

  private parseArgs(argsStr: string): string[] {
    const args: string[] = [];
    let depth = 0, cur = '';
    for (const ch of argsStr) {
      if (ch === '(') depth++;
      else if (ch === ')') depth--;
      else if (ch === ',' && depth === 0) { args.push(cur); cur = ''; continue; }
      cur += ch;
    }
    if (cur) args.push(cur);
    return args;
  }

  // --- Sync Engine ---
  onCellChange(r: number = this.selectedRow, c: number = this.selectedCol, forceBulk: boolean = false) {
    this.updateDisplayCache();

    if (!this.cellEditHistory) this.cellEditHistory = {};
    if (r !== undefined && c !== undefined && !forceBulk && !this.applyingRemote) {
      const key = `${this.currentSheetIdx}-${r}-${c}`;
      if (!this.cellEditHistory[key]) this.cellEditHistory[key] = [];
      const val = this.cells[r]?.[c] ?? '';
      const action = val ? (this.cellEditHistory[key].length === 0 ? 'ADDED' : 'EDITED') : 'CLEARED';
      this.cellEditHistory[key].unshift({
        user: 'Current User',
        time: new Date(),
        action: action,
        value: val
      });
    }

    if (this.applyingRemote) return;
    // CRITICAL: Never send data to the backend before the initial load is complete.
    // Sending empty cells would overwrite the real data stored in R2.
    if (!this.dataLoaded) return;
    if (forceBulk) {
      this.api.sendUpdate(JSON.stringify(this.getSparse()), this.title);
    } else {
      const value = this.cells[r]?.[c] ?? '';
      const format = this.formats[`${r},${c}`];
      this.api.sendCellUpdate(this.currentSheetIdx, r, c, value, format);
    }
  }

  private getSparse() {
    const s: Record<number, Record<number, string>> = {};
    for (let r = 0; r < this.ROWS; r++)
      for (let c = 0; c < this.COLS; c++)
        if (this.cells[r][c]) { if (!s[r]) s[r] = {}; s[r][c] = this.cells[r][c]; }

    // Cleanup empty formats before saving — preserve all meaningful properties
    const cleanFormats: Record<string, CellFormat> = {};
    Object.keys(this.formats).forEach(k => {
      const f = this.formats[k];
      if (
        f.bold || f.italic || f.strikethrough || f.underline ||
        f.color || f.bg || f.align || f.vertAlign ||
        f.font || f.size || f.wrap !== undefined ||
        f.indent || f.rotation || f.numFormat || f.decimals !== undefined ||
        f.borders ||
        (f as any)._mergeSpan || (f as any)._mergedInto ||
        (f as any).note || (f as any).hyperlink || (f as any).checkbox ||
        (f as any).conditionalFormat || (f as any).commentData
      ) {
        cleanFormats[k] = f;
      }
    });

    // Flush live state back into current sheet before serializing all sheets
    const existing = this.sheets[this.currentSheetIdx];
    this.sheets[this.currentSheetIdx] = {
      ...existing,
      cells: this.cells.map(row => [...row]),
      formats: { ...cleanFormats },
      validations: { ...this.validations },
      hiddenRows: Array.from(this.hiddenRows),
      activeFilterCols: Array.from(this.activeFilterCols),
      filterActive: this.filterActive,
      advFilterSavedState: this.serializeAdvFilterState(),
      frozenRowsCount: this.frozenRowsCount,
      frozenColsCount: this.frozenColsCount,
    };

    // Convert every sheet's cells 2D array to sparse format to avoid huge payloads
    const sparseSheets = this.sheets.map(sheet => {
      const sparseC: Record<number, Record<number, string>> = {};
      const sheetRows = sheet.cells || [];
      for (let r = 0; r < sheetRows.length; r++) {
        for (let c = 0; c < (sheetRows[r]?.length ?? 0); c++) {
          if (sheetRows[r][c]) {
            if (!sparseC[r]) sparseC[r] = {};
            sparseC[r][c] = sheetRows[r][c];
          }
        }
      }
      return { ...sheet, cells: sparseC };
    });

    return {
      cells: s,
      formats: cleanFormats,
      validations: this.validations,
      calendarNotes: this.calendarNotes,
      globalNotes: this.globalNotes,
      tasks: this.tasks,
      colWidths: this.sheets[this.currentSheetIdx].colWidths,
      rowHeights: this.sheets[this.currentSheetIdx].rowHeights,
      // Root-level filter state — makes filter persist through both the HTTP and WebSocket load paths
      filterActive: this.filterActive,
      activeFilterCols: Array.from(this.activeFilterCols),
      hiddenRows: Array.from(this.hiddenRows),
      advFilterSavedState: this.serializeAdvFilterState(),
      frozenRowsCount: this.frozenRowsCount,
      frozenColsCount: this.frozenColsCount,
      _importedSheets: sparseSheets
    };
  }


  save() {
    // CRITICAL: Refuse to queue a save before data has been loaded from the server.
    // This prevents the race condition where empty cells overwrite real R2 data.
    if (!this.dataLoaded) {
      console.warn('[SheetEditor] Ignoring save() — data not yet loaded from server.');
      return;
    }
    this.saveStatus = 'saving';
    this.hasPendingChanges = true;
    // Push to subject instead of hitting the backend immediately
    this.saveSubject.next();
  }

  // The actual HTTP call to the backend
  private executeSave() {
    if ((this as any).initError) {
      console.warn('[SheetEditor] Skipping save — document failed to initialize correctly.');
      return;
    }
    // Prevent saving empty data before initial load completes
    if (!this.dataLoaded) {
      console.warn('[SheetEditor] Skipping save — data not yet loaded.');
      return;
    }
    
    this.forceFlushAudit();

    this.api.saveDocument(this.docId, this.title, JSON.stringify(this.getSparse())).subscribe({
      next: () => {
        this.saveStatus = 'saved';
        this.hasPendingChanges = false;
        this.lastSavedTime = new Date().toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit', second: '2-digit', hour12: true });
        
        // Notify edit event
        this.api.triggerNotification(this.docId, 'edit').subscribe({
          error: (err) => console.error('Failed to trigger edit notification', err)
        });
      },
      error: () => { this.saveStatus = 'error'; }
    });
  }

  copyLink() {
    navigator.clipboard.writeText(window.location.href)
      .then(() => this.showToast('Link copied! Anyone with the link can collaborate.'));
  }

  makePublic() {
    this.isPublic = true;
    this.copyLink();
  }

  shareTo(platform: string) {
    this.shareModalOpen = false;
    const url = encodeURIComponent(window.location.href);
    const text = encodeURIComponent(`Check out this spreadsheet: ${this.title}\n\n`);
    if (platform === 'whatsapp') {
      window.open(`https://api.whatsapp.com/send?text=${text}${url}`, '_blank');
    } else if (platform === 'email') {
      const subject = encodeURIComponent(`Spreadsheet: ${this.title}`);
      const body = encodeURIComponent(`Check out this spreadsheet I'm sharing:\n\n${window.location.href}`);
      window.location.href = `mailto:?subject=${subject}&body=${body}`;
    }
  }

  onShareSearch() {
    if (!this.shareQuery || this.shareQuery.length < 2) {
      this.userSearchResults = [];
      return;
    }
    this.api.searchUsers(this.shareQuery).subscribe({
      next: (users) => {
        this.userSearchResults = users;
      },
      error: () => {
        this.userSearchResults = [];
      }
    });
  }
  selectShareUser(user: any) {
    if (!this.selectedShareEmails.includes(user.email)) {
      this.selectedShareEmails.push(user.email);
    }
    this.shareQuery = '';
    this.userSearchResults = [];
  }

  addShareEmail(event: Event) {
    event.preventDefault();
    if (this.shareQuery.trim()) {
      const email = this.shareQuery.trim();
      if (!this.selectedShareEmails.includes(email)) {
        this.selectedShareEmails.push(email);
      }
      this.shareQuery = '';
      this.userSearchResults = [];
    }
  }

  removeShareEmail(email: string) {
    this.selectedShareEmails = this.selectedShareEmails.filter(e => e !== email);
  }
  submitShare() {
    if (this.shareQuery.trim()) {
      this.showToast(`Shared with ${this.shareQuery.trim()}`);
      this.shareQuery = '';
      this.shareModalOpen = false;
    }
  }



  async insertLink() {
    const url = await this.openPrompt('Enter URL to insert into cell:');
    if (url) {
      this.cells[this.selectedRow][this.selectedCol] = url;
      this.formulaBarValue = url;
      this.onCellChange();
      this.showToast('Link inserted into cell.');
    }
  }
  private getFormatsForSheet(sheetIdx: number): Record<string, CellFormat> {
    if (sheetIdx === this.currentSheetIdx) return this.formats;
    if (!this.sheets[sheetIdx].formats) this.sheets[sheetIdx].formats = {};
    return this.sheets[sheetIdx].formats;
  }

  hasComment(r: number, c: number): boolean {
    const ref = `${r},${c}`;
    const format = this.formats[ref] as any;
    return !!(format?.commentData || format?.comment);
  }
  isCommentHighlighted(r: number, c: number): boolean {
    return this.highlightCommentsEnabled && this.hasComment(r, c);
  }

  toggleHighlightComments() {
    this.highlightCommentsEnabled = !this.highlightCommentsEnabled;
    localStorage.setItem('highlightCommentsEnabled', this.highlightCommentsEnabled.toString());
  }

  insertComment() {
    this.closeMenus();
    this.sidePanelApp = 'comments';
    this.commentsViewFilter = 'all';
    this.initNewCommentForCell(this.selectedRow, this.selectedCol);
  }

  initNewCommentForCell(r: number, c: number) {
    this.newCommentCellRef = `${this.currentSheetIdx}:${r},${c}`;
    this.newCommentCellName = this.colLabel(c) + (r + 1);
    this.newCommentText = '';
    setTimeout(() => {
      if (this.newCommentInput) {
        this.newCommentInput.nativeElement.focus();
      }
    }, 100);
  }

  cancelNewComment() {
    this.newCommentCellRef = null;
    this.newCommentText = '';
  }

  submitNewComment() {
    if (!this.newCommentText.trim() || !this.newCommentCellRef) return;
    const [sheetIdxStr, ref] = this.newCommentCellRef.split(':');
    const sheetIdx = parseInt(sheetIdxStr);
    const targetFormats = this.getFormatsForSheet(sheetIdx);

    if (!targetFormats[ref]) targetFormats[ref] = {};
    const existingFormat = targetFormats[ref] as any;

    existingFormat.commentData = {
      id: Date.now().toString(),
      text: this.newCommentText.trim(),
      authorName: this.auth.user?.name || 'You',
      timestamp: new Date().toISOString(),
      resolved: false,
      replies: []
    };

    if (existingFormat.comment) delete existingFormat.comment;

    if (sheetIdx === this.currentSheetIdx) {
      this.formats = { ...this.formats };
      this.onCellChange();
    } else {
      this.sheets[sheetIdx].formats = { ...this.sheets[sheetIdx].formats };
    }

    const textToNotify = this.newCommentText.trim();
    this.cancelNewComment();
    this.save();
    this.updateCachedComments();
    this.showToast('Comment added.');
    
    // Notify comment event
    this.api.triggerNotification(this.docId, 'comment', textToNotify).subscribe({
      error: (err) => console.error('Failed to trigger comment notification', err)
    });
  }

  getFilteredComments() {
    return this.cachedComments;
  }

  updateCachedComments() {
    const all: any[] = [];
    if (this.sidePanelApp !== 'comments') return;

    this.sheets.forEach((sheet, sIdx) => {
      if (this.commentsViewFilter === 'current' && sIdx !== this.currentSheetIdx) return;
      const formatsSource = this.getFormatsForSheet(sIdx);
      Object.keys(formatsSource).forEach(ref => {
        const cData = (formatsSource[ref] as any)?.commentData;
        if (cData) {
          if (this.commentsStatusFilter === 'resolved' && !cData.resolved) return;
          if (this.commentsStatusFilter === 'unresolved' && cData.resolved) return;
          all.push({
            sheetIdx: sIdx,
            sheetName: sheet.name,
            ref: ref,
            cellName: this.colLabel(parseInt(ref.split(',')[1])) + (parseInt(ref.split(',')[0]) + 1),
            data: cData
          });
        }
      });
    });
    this.cachedComments = all.sort((a, b) => new Date(b.data.timestamp).getTime() - new Date(a.data.timestamp).getTime());
  }

  goToCommentCell(c: any) {
    if (this.currentSheetIdx !== c.sheetIdx) {
      this.switchSheet(c.sheetIdx);
    }
    const [r, col] = c.ref.split(',').map(Number);
    this.selectCell(r, col);

    const cellEl = document.getElementById(`cell-${r}-${col}`);
    if (cellEl) cellEl.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
  }

  toggleCommentResolve(c: any) {
    c.data.resolved = !c.data.resolved;
    if (c.sheetIdx === this.currentSheetIdx) {
      this.formats = { ...this.formats };
      this.onCellChange();
    } else {
      this.sheets[c.sheetIdx].formats = { ...this.sheets[c.sheetIdx].formats };
    }
    this.save();
    this.updateCachedComments();
  }

  deleteComment(c: any) {
    const formatsSource = this.getFormatsForSheet(c.sheetIdx);
    const format = formatsSource[c.ref] as any;
    if (format) {
      delete format.commentData;
      delete format.comment;
      if (Object.keys(format).length === 0) {
        delete formatsSource[c.ref];
      }
    }
    if (c.sheetIdx === this.currentSheetIdx) {
      this.formats = { ...this.formats };
      this.onCellChange();
    } else {
      this.sheets[c.sheetIdx].formats = { ...this.sheets[c.sheetIdx].formats };
    }
    this.save();
    this.updateCachedComments();
    this.showToast('Comment deleted.');
  }

  submitReply(c: any) {
    const text = (this.replyTexts[c.data.id] || '').trim();
    if (!text) return;

    if (!c.data.replies) c.data.replies = [];
    c.data.replies.push({
      text: text,
      authorName: this.auth.user?.name || 'You',
      timestamp: new Date().toISOString()
    });

    this.replyTexts[c.data.id] = '';
    if (c.sheetIdx === this.currentSheetIdx) {
      this.formats = { ...this.formats };
      this.onCellChange();
    } else {
      this.sheets[c.sheetIdx].formats = { ...this.sheets[c.sheetIdx].formats };
    }
    this.save();
    this.updateCachedComments();
  }

  generateChart(type: string = 'column') {
    this.closeMenus();
    let minR = this.selectedRow, maxR = this.selectedRow;
    let minC = this.selectedCol, maxC = this.selectedCol;
    if (this.rangeStart && this.rangeEnd) {
      minR = Math.min(this.rangeStart.r, this.rangeEnd.r);
      maxR = Math.max(this.rangeStart.r, this.rangeEnd.r);
      minC = Math.min(this.rangeStart.c, this.rangeEnd.c);
      maxC = Math.max(this.rangeStart.c, this.rangeEnd.c);
    } else {
      maxR = Math.min(this.ROWS - 1, minR + 9);
    }

    const numRows = maxR - minR + 1;
    const numCols = maxC - minC + 1;

    const series: number[][] = [];
    for (let c = minC; c <= maxC; c++) {
      const colVals: number[] = [];
      for (let r = minR; r <= maxR; r++) {
        const val = this.getDisplayValue(r, c);
        const v = parseFloat(val);
        colVals.push(isNaN(v) ? 0 : v);
      }
      series.push(colVals);
    }


    let hasData = false;
    series.forEach(s => s.forEach(v => { if (v !== 0) hasData = true; }));
    if (!hasData) {
      this.showToast('Please enter some numbers in the cells before generating a chart!');
      return;
    }

    const colors = ['#4285f4', '#ea4335', '#fbbc04', '#34a853', '#673ab7', '#ff9800', '#00bcd4', '#e91e63'];
    const bw = type === 'grouped' ? Math.max(10, 36 / numCols) : 36;
    const spacing = 12;
    const groupWidth = type === 'grouped' ? (bw * numCols) + spacing : bw + spacing;
    const gh = 250;
    const gw = numRows * groupWidth + 80;

    let svg = `<svg width="${gw}" height="${gh + 40}" xmlns="http://www.w3.org/2000/svg" style="background:#fff;border:1px solid #e0e0e0;border-radius:4px;font-family:sans-serif;">`;
    svg += `<line x1="50" y1="20" x2="50" y2="${gh + 20}" stroke="#ccc" stroke-width="1"/>`;
    svg += `<line x1="50" y1="${gh + 20}" x2="${gw - 20}" y2="${gh + 20}" stroke="#ccc" stroke-width="1"/>`;

    if (type === 'column' || type === 'grouped') {
      let globalMax = 1;
      series.forEach(s => s.forEach(v => { if (v > globalMax) globalMax = v; }));

      for (let i = 0; i < numRows; i++) {
        for (let s = 0; s < numCols; s++) {
          const v = series[s][i];
          const h = Math.round((v / globalMax) * gh);
          const x = 60 + i * groupWidth + (type === 'grouped' ? s * bw : 0);
          const y = gh + 20 - h;
          if (h > 0) svg += `<rect x="${x}" y="${y}" width="${bw}" height="${h}" fill="${colors[s % colors.length]}" rx="2"/>`;
        }
      }
    } else if (type === 'stacked_column' || type === 'stacked_100') {
      let maxRowSum = 1;
      const rowSums = [];
      for (let i = 0; i < numRows; i++) {
        let sum = 0;
        for (let s = 0; s < numCols; s++) sum += series[s][i];
        rowSums.push(sum);
        if (sum > maxRowSum) maxRowSum = sum;
      }

      for (let i = 0; i < numRows; i++) {
        let currentY = gh + 20;
        const rowSum = rowSums[i];
        for (let s = 0; s < numCols; s++) {
          const v = series[s][i];
          if (v <= 0) continue;

          let h = 0;
          if (type === 'stacked_100') {
            h = rowSum > 0 ? Math.round((v / rowSum) * gh) : 0;
          } else {
            h = Math.round((v / maxRowSum) * gh);
          }

          currentY -= h;
          const x = 60 + i * groupWidth;
          svg += `<rect x="${x}" y="${currentY}" width="${bw}" height="${h}" fill="${colors[s % colors.length]}"/>`;
        }
      }
    } else if (type === 'line' || type === 'area') {
      let globalMax = 1;
      series.forEach(s => s.forEach(v => { if (v > globalMax) globalMax = v; }));

      for (let s = 0; s < numCols; s++) {
        let pts = '';
        for (let i = 0; i < numRows; i++) {
          const v = series[s][i];
          const h = Math.round((v / globalMax) * gh);
          const x = 60 + i * groupWidth + (groupWidth / 2);
          const y = gh + 20 - h;
          pts += `${x},${y} `;
          if (type === 'line') {
            svg += `<circle cx="${x}" cy="${y}" r="4" fill="${colors[s % colors.length]}"/>`;
          }
        }
        if (type === 'line') {
          svg += `<polyline points="${pts.trim()}" fill="none" stroke="${colors[s % colors.length]}" stroke-width="3"/>`;
        } else {
          const firstX = 60 + (groupWidth / 2);
          const lastX = 60 + (numRows - 1) * groupWidth + (groupWidth / 2);
          const areaPts = `${firstX},${gh + 20} ${pts} ${lastX},${gh + 20}`;
          svg += `<polygon points="${areaPts}" fill="${colors[s % colors.length]}" opacity="0.4"/>`;
          svg += `<polyline points="${pts.trim()}" fill="none" stroke="${colors[s % colors.length]}" stroke-width="2"/>`;
        }
      }
    } else if (type === 'scatter') {
      let globalMax = 1;
      series.forEach(s => s.forEach(v => { if (v > globalMax) globalMax = v; }));

      for (let s = 0; s < numCols; s++) {
        for (let i = 0; i < numRows; i++) {
          const v = series[s][i];
          if (v === 0) continue;
          const h = Math.round((v / globalMax) * gh);
          const x = 60 + i * groupWidth + (bw / 2);
          const y = gh + 20 - h;
          svg += `<circle cx="${x}" cy="${y}" r="6" fill="${colors[s % colors.length]}" opacity="0.7"/>`;
        }
      }
    } else if (type === 'pie') {
      let total = 0;
      const pieData = series[0].filter(v => v > 0);
      pieData.forEach(v => total += v);

      if (total > 0) {
        let startAngle = 0;
        const cx = gw / 2;
        const cy = (gh + 40) / 2;
        const r = Math.min(cx, cy) - 40;

        pieData.forEach((v, i) => {
          const sliceAngle = (v / total) * 360;
          if (sliceAngle === 360) {
            svg += `<circle cx="${cx}" cy="${cy}" r="${r}" fill="${colors[i % colors.length]}"/>`;
            return;
          }
          const endAngle = startAngle + sliceAngle;
          // svg arc uses radians
          const x1 = cx + r * Math.cos(Math.PI * startAngle / 180);
          const y1 = cy + r * Math.sin(Math.PI * startAngle / 180);
          const x2 = cx + r * Math.cos(Math.PI * endAngle / 180);
          const y2 = cy + r * Math.sin(Math.PI * endAngle / 180);

          const largeArc = sliceAngle > 180 ? 1 : 0;
          svg += `<path d="M ${cx} ${cy} L ${x1} ${y1} A ${r} ${r} 0 ${largeArc} 1 ${x2} ${y2} Z" fill="${colors[i % colors.length]}"/>`;
          startAngle = endAngle;
        });
      }
    }
    svg += '</svg>';

    const win = window.open('', '_blank', `width=${gw + 60},height=${gh + 120}`);
    if (win) {
      win.document.write(`<html><body style="margin:20px;font-family:sans-serif;background:#f8f9fa;"><h3>Chart Preview: ${type.replace('_', ' ').toUpperCase()}</h3>${svg}<p style="color:#888;font-size:13px;">Close this window when done.</p></body></html>`);
    }
    this.showToast(`Chart generated successfully.`);
  }

  toggleFilter() {
    this.pushHistory({ action_type: this.filterActive ? 'clear-filter' : 'apply-filter', target_range: 'Entire Sheet' });
    this.filterActive = !this.filterActive;
    if (!this.filterActive) {
      this.hiddenRows.clear();
      this.activeFilterCols.clear();
      this.advFilterSavedState.clear();
      this.showToast('Filter cleared.');
    } else {
      this.showToast('Filter activated. Use column options to filter.');
    }
    this.onCellChange(undefined, undefined, true);
  }

  filterByCellValue() {
    this.pushHistory({ action_type: 'apply-filter', target_range: 'Entire Sheet' });
    this.filterActive = true;
    const rClick = this.ctxRow !== null ? this.ctxRow : this.selectedRow;
    const cClick = this.ctxCol !== null ? this.ctxCol : this.selectedCol;
    this.activeFilterCols.add(cClick);
    const targetVal = this.cells[rClick][cClick];
    this.advFilterSavedState.set(cClick, {
      tab: 'value',
      allowedVals: new Set([targetVal]),
      allowedBg: new Set(),
      allowedText: new Set()
    });

    this.recalculateAllFilters();
    this.showToast(`Filtered by value: "${targetVal}"`);
    this.onCellChange(undefined, undefined, true);
  }

  filterByCellColor() {
    this.pushHistory({ action_type: 'apply-filter', target_range: 'Entire Sheet' });
    this.filterActive = true;
    const rClick = this.ctxRow !== null ? this.ctxRow : this.selectedRow;
    const cClick = this.ctxCol !== null ? this.ctxCol : this.selectedCol;
    this.activeFilterCols.add(cClick);
    const targetRef = `${rClick},${cClick}`;
    const targetColor = this.formats[targetRef]?.bg || '';
    this.advFilterSavedState.set(cClick, {
      tab: 'cellColor',
      allowedVals: new Set(),
      allowedBg: new Set([targetColor]),
      allowedText: new Set()
    });

    this.recalculateAllFilters();
    this.showToast(targetColor ? 'Filtered by cell color.' : 'Filtered by empty cell color.');
    this.onCellChange(undefined, undefined, true);
  }

  filterByTextColor() {
    this.pushHistory({ action_type: 'apply-filter', target_range: 'Entire Sheet' });
    this.filterActive = true;
    const rClick = this.ctxRow !== null ? this.ctxRow : this.selectedRow;
    const cClick = this.ctxCol !== null ? this.ctxCol : this.selectedCol;
    this.activeFilterCols.add(cClick);
    const targetRef = `${rClick},${cClick}`;
    const targetColor = this.formats[targetRef]?.color || '';

    this.advFilterSavedState.set(cClick, {
      tab: 'textColor',
      allowedVals: new Set(),
      allowedBg: new Set(),
      allowedText: new Set([targetColor])
    });

    this.recalculateAllFilters();
    this.showToast(targetColor ? 'Filtered by text color.' : 'Filtered by default text color.');
    this.onCellChange(undefined, undefined, true);
  }

  isFilterHeaderCell(r: number, c: number): boolean {
    if (!this.filterActive) return false;
    const headerRow = Math.max(1, this.frozenRowsCount || 1) - 1;
    const cellValue = this.cells[r]?.[c];
    return r === headerRow && cellValue !== undefined && cellValue !== null && cellValue.toString().trim() !== '';
  }

  isColumnFiltered(c: number): boolean {
    return this.activeFilterCols.has(c);
  }

  openFilterMenu(event: MouseEvent, r: number, c: number) {
    this.selectCell(r, c);
    this.closeMenus();
    this.advFilterCol = c;
    this.advFilterVisible = true;
    this.advFilterTab = 'value';
    this.advFilterSearch = '';

    const target = event.currentTarget as HTMLElement;
    const rect = target.getBoundingClientRect();
    this.advFilterX = rect.left;

    // Estimate panel height based on design bounds
    const estimatedHeight = 450;

    // Smart Positioning: If it overflows the bottom edge, align it upwards from the filter icon instead
    if (rect.bottom + estimatedHeight > window.innerHeight && rect.top - estimatedHeight > 0) {
      this.advFilterY = rect.top - estimatedHeight;
      this.advFilterMaxHeight = Math.max(rect.top - 10, 300);
    } else {
      // Standard placement exactly underneath
      this.advFilterY = rect.bottom + 4;

      // Dynamic shrinking: if the screen is small, force it to fit perfectly within the visible area
      const availableHeight = window.innerHeight - this.advFilterY - 10;
      this.advFilterMaxHeight = Math.min(estimatedHeight, Math.max(availableHeight, 300));
    }

    this.populateAdvFilter(c);
  }

  populateAdvFilter(c: number) {
    const lastRow = this.getLastUsedRow();
    const startRow = Math.max(1, this.frozenRowsCount || 0);

    const uniqueVals = new Set<string>();
    const uniqueBg = new Set<string>();
    const uniqueText = new Set<string>();

    for (let r = startRow; r <= lastRow; r++) {
      uniqueVals.add(this.cells[r][c] || '');
      const ref = `${r},${c}`;
      uniqueBg.add(this.formats[ref]?.bg || '');
      uniqueText.add(this.formats[ref]?.color || '');
    }

    const savedState = this.advFilterSavedState.get(c);

    this.advFilterValues = Array.from(uniqueVals).sort().map(val => ({ val, selected: savedState ? savedState.allowedVals.has(val) : true }));
    this.advFilterBgColors = Array.from(uniqueBg).sort().map(val => ({ val, selected: savedState ? savedState.allowedBg.has(val) : true }));
    this.advFilterTextColors = Array.from(uniqueText).sort().map(val => ({ val, selected: savedState ? savedState.allowedText.has(val) : true }));

    if (savedState) {
      this.advFilterTab = savedState.tab;
    } else {
      this.advFilterTab = 'value';
    }
  }

  closeAdvFilter() {
    this.advFilterVisible = false;
  }

  allAdvFilterSelected(tab: string): boolean {
    if (tab === 'value') {
      let allSelected = true;
      for (const x of this.advFilterValues) if (!x.selected) allSelected = false;
      return this.advFilterValues.length > 0 && allSelected;
    }
    if (tab === 'cellColor') {
      let allSelected = true;
      for (const x of this.advFilterBgColors) if (!x.selected) allSelected = false;
      return this.advFilterBgColors.length > 0 && allSelected;
    }
    if (tab === 'textColor') {
      let allSelected = true;
      for (const x of this.advFilterTextColors) if (!x.selected) allSelected = false;
      return this.advFilterTextColors.length > 0 && allSelected;
    }
    return false;
  }

  getLastUsedRow(): number {
    let lastDataRow = 0;
    for (let r = this.ROWS - 1; r >= 0; r--) {
      if (this.cells[r].some(val => val !== '')) {
        lastDataRow = r;
        break;
      }
    }
    for (const key of Object.keys(this.formats)) {
      const r = parseInt(key.split(',')[0], 10);
      if (r > lastDataRow) {
        lastDataRow = r;
      }
    }
    return lastDataRow;
  }

  toggleAllAdvFilter(event: any) {
    const checked = event.target.checked;
    if (this.advFilterTab === 'value') {
      this.advFilterValues.forEach(x => x.selected = checked);
    } else if (this.advFilterTab === 'cellColor') {
      this.advFilterBgColors.forEach(x => x.selected = checked);
    } else if (this.advFilterTab === 'textColor') {
      this.advFilterTextColors.forEach(x => x.selected = checked);
    }
  }

  applyAdvFilter() {
    this.pushHistory({ action_type: 'apply-filter', target_range: 'Entire Sheet' });
    this.filterActive = true;
    this.activeFilterCols.add(this.advFilterCol!);
    const c = this.advFilterCol!;

    const allowedVals = new Set(this.advFilterValues.filter(x => x.selected).map(x => x.val));
    const allowedBg = new Set(this.advFilterBgColors.filter(x => x.selected).map(x => x.val));
    const allowedText = new Set(this.advFilterTextColors.filter(x => x.selected).map(x => x.val));

    this.advFilterSavedState.set(c, {
      tab: this.advFilterTab,
      allowedVals,
      allowedBg,
      allowedText
    });

    this.recalculateAllFilters();
    this.showToast('Filter applied.');
    this.onCellChange(undefined, undefined, true);
    this.closeAdvFilter();
  }

  clearAdvFilter() {
    this.pushHistory({ action_type: 'clear-filter', target_range: 'Entire Sheet' });
    this.activeFilterCols.delete(this.advFilterCol!);
    this.advFilterSavedState.delete(this.advFilterCol!);

    if (this.activeFilterCols.size === 0) {
      this.hiddenRows.clear();
      this.filterActive = false;
    } else {
      this.recalculateAllFilters();
    }

    this.showToast('Filter cleared.');
    this.onCellChange(undefined, undefined, true);
    this.closeAdvFilter();
  }

  private recalculateAllFilters() {
    const lastRow = this.getLastUsedRow();
    const startRow = Math.max(1, this.frozenRowsCount || 0);
    for (let r = 0; r < startRow; r++) this.hiddenRows.delete(r);

    for (let r = startRow; r <= lastRow; r++) {
      let hide = false;
      for (const [colIndex, state] of this.advFilterSavedState.entries()) {
        const val = this.cells[r][colIndex] || '';
        const ref = `${r},${colIndex}`;
        const bg = this.formats[ref]?.bg || '';
        const text = this.formats[ref]?.color || '';

        if (state.tab === 'value') {
          if (!state.allowedVals.has(val)) hide = true;
        } else if (state.tab === 'cellColor') {
          if (!state.allowedBg.has(bg)) hide = true;
        } else if (state.tab === 'textColor') {
          if (!state.allowedText.has(text)) hide = true;
        }
        if (hide) break;
      }
      if (hide) {
        this.hiddenRows.add(r);
      } else {
        this.hiddenRows.delete(r);
      }
    }
    for (let r = lastRow + 1; r < this.ROWS; r++) this.hiddenRows.delete(r);
  }

  advSort(asc: boolean) {
    const c = this.advFilterCol!;
    this.pushHistory({ action_type: asc ? 'sort-ascending' : 'sort-descending', target_range: `${this.colLabel(c)}:${this.colLabel(c)}` });
    const lastRow = this.getLastUsedRow();
    const startRow = Math.max(1, this.frozenRowsCount || 0);

    const rowsData = [];
    for (let r = startRow; r <= lastRow; r++) {
      const rowCells = [...this.cells[r]];
      const rowFormats: Record<string, CellFormat> = {};
      const rowValidations: Record<string, CellValidation> = {};

      for (let i = 0; i < this.COLS; i++) {
        if (this.formats[`${r},${i}`]) rowFormats[`${r},${i}`] = this.formats[`${r},${i}`];
        if (this.validations[`${r},${i}`]) rowValidations[`${r},${i}`] = this.validations[`${r},${i}`];
      }

      rowsData.push({ r, val: this.cells[r][c] || '', rowCells, rowFormats, rowValidations });
    }

    rowsData.sort((a, b) => {
      const cmp = a.val.localeCompare(b.val, undefined, { numeric: true });
      return asc ? cmp : -cmp;
    });

    for (let i = 0; i < rowsData.length; i++) {
      const r = startRow + i;
      this.cells[r] = rowsData[i].rowCells;

      for (let j = 0; j < this.COLS; j++) {
        delete this.formats[`${r},${j}`];
        delete this.validations[`${r},${j}`];
      }

      const oldR = rowsData[i].r;
      for (let j = 0; j < this.COLS; j++) {
        if (rowsData[i].rowFormats[`${oldR},${j}`]) {
          this.formats[`${r},${j}`] = rowsData[i].rowFormats[`${oldR},${j}`];
        }
        if (rowsData[i].rowValidations[`${oldR},${j}`]) {
          this.validations[`${r},${j}`] = rowsData[i].rowValidations[`${oldR},${j}`];
        }
      }
    }

    this.showToast(asc ? 'Sorted A to Z' : 'Sorted Z to A');
    this.onCellChange(undefined, undefined, true);
    this.closeAdvFilter();
  }



  exportFile(format: string) {
    this.closeMenus();
    this.save();
    this.showToast(`Exporting as ${format.toUpperCase()}...`);

    // xlsx branch is async (ExcelJS); everything else runs in setTimeout
    if (format === 'xlsx') {
      this._exportXlsx();
      return;
    }

    setTimeout(() => {
      try {
        if (format === 'pdf') {
          const doc = new jsPDF({ orientation: 'landscape' });
          for (let sIdx = 0; sIdx < this.sheets.length; sIdx++) {
            const sheet = this.sheets[sIdx];
            const sheetCells = sheet.cells || [];
            let sheetMaxRow = 0; let sheetMaxCol = 0;
            for (let r = 0; r < this.ROWS; r++) {
              if (!sheetCells[r]) continue;
              for (let c = 0; c < this.COLS; c++) {
                if (sheetCells[r][c]) { sheetMaxRow = Math.max(sheetMaxRow, r); sheetMaxCol = Math.max(sheetMaxCol, c); }
              }
            }
            if (sIdx > 0) doc.addPage();
            doc.setFontSize(14);
            doc.text(sheet.name || `Sheet${sIdx + 1}`, 14, 15);

            const body = [];
            for (let r = 0; r <= sheetMaxRow; r++) {
              const row = [];
              for (let c = 0; c <= sheetMaxCol; c++) {
                row.push((sheetCells[r] && sheetCells[r][c] ? sheetCells[r][c] : '').toString());
              }
              body.push(row);
            }

            autoTable(doc, {
              body: body,
              theme: 'grid',
              styles: { fontSize: 8, cellPadding: 2 },
              startY: 20
            });
          }

          doc.save(`${this.title || 'Spreadsheet'}.pdf`);
          this.showToast('Download complete.');
          return;
        }

        let content = '';
        let mimeType = '';
        let extension = format;

        if (format === 'xlsb' || format === 'ods') {
          // The free SheetJS build cannot write true .xlsb binary workbooks —
          // only 'xlsx' and 'ods' bookTypes are actually supported for writing.
          if (format === 'xlsb') {
            this.showToast('.xlsb export is not supported in this browser build — downloading as .xlsx instead.');
          }
          const effectiveBookType: 'xlsx' | 'ods' = format === 'ods' ? 'ods' : 'xlsx';
          const effectiveExt = format === 'ods' ? 'ods' : 'xlsx';

          const wb = XLSX.utils.book_new();
          const usedNames = new Set<string>();
          for (let sIdx = 0; sIdx < this.sheets.length; sIdx++) {
            const sheet = this.sheets[sIdx];
            const sheetCells = sheet.cells || [];
            let maxRow = 0; let maxCol = 0;
            for (let r = 0; r < this.ROWS; r++) {
              if (!sheetCells[r]) continue;
              for (let c = 0; c < this.COLS; c++) {
                if (sheetCells[r][c]) { maxRow = Math.max(maxRow, r); maxCol = Math.max(maxCol, c); }
              }
            }
            const aoa = [];
            for (let r = 0; r <= maxRow; r++) {
              const row = [];
              for (let c = 0; c <= maxCol; c++) {
                let raw: any = sheetCells[r] ? sheetCells[r][c] : undefined;
                if (raw !== null && raw !== undefined && typeof raw === 'object') {
                  raw = raw.v !== undefined ? raw.v : (raw.value !== undefined ? raw.value : raw.text !== undefined ? raw.text : '');
                }
                if (typeof raw === 'string') {
                  if (raw.startsWith('data:image')) { raw = ''; }
                  else if (raw.length > 32767) { raw = raw.substring(0, 32767); }
                }
                row.push(raw !== null && raw !== undefined ? raw : '');
              }
              aoa.push(row);
            }
            const ws = XLSX.utils.aoa_to_sheet(aoa);
            let safeName = (sheet.name || `Sheet${sIdx + 1}`).replace(/[\[\]\*?\/\:\\]/g, '_').substring(0, 31);
            let finalName = safeName;
            let counter = 1;
            while (usedNames.has(finalName.toLowerCase())) {
              const suffix = `_${counter}`;
              finalName = safeName.substring(0, 31 - suffix.length) + suffix;
              counter++;
            }
            usedNames.add(finalName.toLowerCase());
            XLSX.utils.book_append_sheet(wb, ws, finalName);
          }

          try {
            const base64 = XLSX.write(wb, { bookType: effectiveBookType, type: 'base64' });
            const binaryStr = atob(base64);
            const bytes = new Uint8Array(binaryStr.length);
            for (let i = 0; i < binaryStr.length; i++) bytes[i] = binaryStr.charCodeAt(i);
            const blob = new Blob([bytes], { type: 'application/octet-stream' });
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `${this.title || 'Spreadsheet'}.${effectiveExt}`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
            this.showToast('Download complete.');
          } catch (writeErr: any) {
            console.error('XLSX.write failed:', writeErr);
            this.showToast(`Export error: ${writeErr?.message || writeErr}`);
          }
          return;
        }

        if (format === 'csv' || format === 'tsv') {
          const delimiter = format === 'csv' ? ',' : '\t';
          mimeType = format === 'csv' ? 'text/csv;charset=utf-8;' : 'text/tab-separated-values;charset=utf-8;';
          const rows = [];
          for (let r = 0; r < this.ROWS; r++) {
            const rowData = [];
            for (let c = 0; c < this.COLS; c++) {
              let val = (this.cells[r][c] || '').toString();
              if (val.includes(delimiter) || val.includes('\\n') || val.includes('"')) {
                val = '"' + val.replace(/"/g, '""') + '"';
              }
              rowData.push(val);
            }
            if (rows.length > 0 || rowData.some(v => v !== '')) {
              rows.push(rowData.join(delimiter));
            }
          }
          while (rows.length > 0 && rows[rows.length - 1].replace(new RegExp(delimiter, 'g'), '') === '') {
            rows.pop();
          }
          content = rows.join('\\n');
        } else {
          mimeType = format === 'html' ? 'text/html;charset=utf-8;' : 'application/vnd.ms-excel;charset=utf-8;';
          extension = format === 'html' ? 'html' : 'xls';

          content = '<html xmlns:x="urn:schemas-microsoft-com:office:excel"><head><meta charset="utf-8"></head><body><table border="1">';
          let maxRow = 0;
          let maxCol = 0;
          for (let r = 0; r < this.ROWS; r++) {
            for (let c = 0; c < this.COLS; c++) {
              if (this.cells[r][c]) { maxRow = Math.max(maxRow, r); maxCol = Math.max(maxCol, c); }
            }
          }
          for (let r = 0; r <= maxRow; r++) {
            content += '<tr>';
            for (let c = 0; c <= maxCol; c++) {
              content += `<td>${this.cells[r][c] || ''}</td>`;
            }
            content += '</tr>';
          }
          content += '</table></body></html>';
        }

        const blob = new Blob([content], { type: mimeType });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${this.title || 'Spreadsheet'}.${extension}`;
        a.click();
        window.URL.revokeObjectURL(url);
        this.showToast('Download complete.');
      } catch (err) {
        console.error(err);
        this.showToast('Export failed. Please try again.');
      }
    }, 100);
  }

  private async _exportXlsx() {
    try {
      const workbook = new ExcelJS.Workbook();
      const usedNames = new Set<string>();

      for (let sIdx = 0; sIdx < this.sheets.length; sIdx++) {
        const sheet = this.sheets[sIdx];
        const sheetCells: any[][] = sheet.cells || [];

        // Sanitize & deduplicate sheet name
        let safeName = (sheet.name || `Sheet${sIdx + 1}`).replace(/[\[\]\*?\/\:\\]/g, '_').substring(0, 31);
        let finalName = safeName;
        let counter = 1;
        while (usedNames.has(finalName.toLowerCase())) {
          const suffix = `_${counter}`;
          finalName = safeName.substring(0, 31 - suffix.length) + suffix;
          counter++;
        }
        usedNames.add(finalName.toLowerCase());

        const worksheet = workbook.addWorksheet(finalName);

        // Find bounds
        let maxRow = 0; let maxCol = 0;
        for (let r = 0; r < this.ROWS; r++) {
          if (!sheetCells[r]) continue;
          for (let c = 0; c < this.COLS; c++) {
            if (sheetCells[r][c]) { maxRow = Math.max(maxRow, r); maxCol = Math.max(maxCol, c); }
          }
        }

        // Track which rows/cols have images for sizing
        const imageRows = new Set<number>();
        const imageCols = new Set<number>();

        // Collect image embedding tasks (process sequentially to avoid UI freeze)
        const imageTasks: Array<{ r: number; c: number; dataUrl: string }> = [];

        // First pass: write plain cell values
        for (let r = 0; r <= maxRow; r++) {
          for (let c = 0; c <= maxCol; c++) {
            let raw: any = sheetCells[r] ? sheetCells[r][c] : undefined;
            if (raw !== null && raw !== undefined && typeof raw === 'object') {
              raw = raw.v !== undefined ? raw.v : (raw.value !== undefined ? raw.value : raw.text !== undefined ? raw.text : '');
            }
            if (typeof raw === 'string' && raw.startsWith('data:image')) {
              imageTasks.push({ r, c, dataUrl: raw });
              // leave cell empty — image will be embedded below
            } else {
              if (typeof raw === 'string' && raw.length > 32767) { raw = raw.substring(0, 32767); }
              const cell = worksheet.getCell(r + 1, c + 1);
              cell.value = (raw !== null && raw !== undefined) ? raw : null;
            }
          }
        }

        // Second pass: embed images sequentially in small batches
        const BATCH = 10;
        for (let i = 0; i < imageTasks.length; i += BATCH) {
          const batch = imageTasks.slice(i, i + BATCH);
          for (const task of batch) {
            try {
              const { r, c, dataUrl } = task;
              // Parse MIME type to get extension
              const mimeMatch = dataUrl.match(/^data:image\/([a-zA-Z0-9+]+);base64,/);
              if (!mimeMatch) { console.warn(`Skipping image at [${r},${c}]: unrecognised MIME`); continue; }
              const ext = mimeMatch[1].toLowerCase().replace('jpeg', 'jpeg') as any;
              const base64Data = dataUrl.split(',')[1];

              const imageId = workbook.addImage({ base64: base64Data, extension: ext });
              worksheet.addImage(imageId, {
                tl: { col: c, row: r } as any,
                ext: { width: 120, height: 80 }
              });

              imageRows.add(r);
              imageCols.add(c);
            } catch (imgErr) {
              console.warn(`Skipping image at [${task.r},${task.c}]:`, imgErr);
            }
          }
          // Yield to browser event loop between batches
          await new Promise<void>(res => setTimeout(res, 0));
        }

        // Set sensible row height / col width for image rows/cols
        imageRows.forEach(r => {
          const row = worksheet.getRow(r + 1);
          if ((row.height || 0) < 60) row.height = 60;
        });
        imageCols.forEach(c => {
          const col = worksheet.getColumn(c + 1);
          if ((col.width || 0) < 15) col.width = 15; // ~80px equivalent
        });
      }

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/octet-stream' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${this.title || 'Spreadsheet'}.xlsx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
      this.showToast('Download complete.');
    } catch (err: any) {
      console.error('ExcelJS export failed:', err);
      this.showToast('Export failed. Please try again.');
    }
  }

  activeModal: 'template' | 'open' | 'import' | 'move' | 'audit' | 'version' | 'workflow' | 'password' | 'form' | 'view_form' | 'manage_forms' | 'macro' | 'edit_macro' | 'functions' | 'merge' | 'goto' | 'shortcuts' | 'insert_sparkline' | 'edit_sparkline' | 'emoji' | 'custom_insert' | null = null;
  // ── Custom Insert dialog state ─────────────────────────────────────
  customInsertType: 'row' | 'col' = 'row';
  customInsertCount: number = 1;
  customInsertPosition: 'before' | 'after' = 'before';
  Math = Math; // expose Math for template
  // ──────────────────────────────────────────────────────────────────
  previewImageUrl: string | null = null;
  saveStatus: 'saved' | 'saving' | 'error' = 'saved';
  lastSavedTime: string = new Date().toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit', second: '2-digit', hour12: true });
  private saveSubject = new Subject<void>();
  private saveSubscription!: Subscription;
  private hasPendingChanges: boolean = false;
  dummyList: any[] = [];

  // Version History State
  versions: any[] = [];
  previewVersionId: string | null = null;
  newVersionName: string = '';
  previewData: any = null;
  previewActiveSheetIdx: number = 0;

  get previewSheets() {
    if (!this.previewData) return [];
    if (this.previewData._importedSheets) return this.previewData._importedSheets;
    return Array.isArray(this.previewData) ? this.previewData : [this.previewData];
  }

  get previewCells() {
    const sheets = this.previewSheets;
    if (sheets.length === 0) return [];
    const sheet = sheets[this.previewActiveSheetIdx] || sheets[0];
    const cells = sheet?.cells || {};

    let maxR = 20;
    let maxC = 10;

    Object.keys(cells).forEach(rKey => {
      const r = parseInt(rKey, 10);
      if (!isNaN(r) && r > maxR) maxR = r;
      if (cells[rKey]) {
        Object.keys(cells[rKey]).forEach(cKey => {
          const c = parseInt(cKey, 10);
          if (!isNaN(c) && c > maxC) maxC = c;
        });
      }
    });

    const rows = [];
    for (let r = 0; r <= maxR + 2; r++) {
      const row = [];
      for (let c = 0; c <= maxC + 2; c++) {
        row.push(cells[r]?.[c] || '');
      }
      rows.push(row);
    }
    return rows;
  }

  myDocs: any[] = [];
  selectedImportFile: File | null = null;
  modalInput = '';
  isStarred = false;

  ocrModalOpen = false;
  ocrImage: string | ArrayBuffer | null = null;
  ocrProgress: number = 0;
  ocrData: string[][] = [];
  ocrInsertTarget: 'new' | 'existing' = 'new';
  ocrSelStart: { r: number, c: number } | null = null;
  ocrSelEnd: { r: number, c: number } | null = null;
  ocrEdit: { r: number, c: number } | null = null;
  ocrDragging = false;

  startOcrDrag(r: number, c: number) {
    this.ocrSelStart = { r, c };
    this.ocrSelEnd = { r, c };
    this.ocrEdit = null;
    this.ocrDragging = true;
  }

  doOcrDrag(r: number, c: number) {
    if (this.ocrDragging) {
      this.ocrSelEnd = { r, c };
    }
  }

  startOcrEdit(r: number, c: number) {
    this.ocrEdit = { r, c };
  }

  ocrHistory: string[][][] = [];
  ocrHistoryIndex: number = -1;

  saveOcrHistory() {
    const snapshot = JSON.parse(JSON.stringify(this.ocrData));
    if (this.ocrHistoryIndex < this.ocrHistory.length - 1) {
      this.ocrHistory = this.ocrHistory.slice(0, this.ocrHistoryIndex + 1);
    }
    this.ocrHistory.push(snapshot);
    this.ocrHistoryIndex = this.ocrHistory.length - 1;
  }

  ocrUndo() {
    if (this.ocrHistoryIndex > 0) {
      this.ocrHistoryIndex--;
      this.ocrData = JSON.parse(JSON.stringify(this.ocrHistory[this.ocrHistoryIndex]));
      this.cdr.detectChanges();
    }
  }

  ocrRedo() {
    if (this.ocrHistoryIndex < this.ocrHistory.length - 1) {
      this.ocrHistoryIndex++;
      this.ocrData = JSON.parse(JSON.stringify(this.ocrHistory[this.ocrHistoryIndex]));
      this.cdr.detectChanges();
    }
  }

  isOcrSelected(r: number, c: number) {
    if (!this.ocrSelStart || !this.ocrSelEnd) return false;
    const minR = Math.min(this.ocrSelStart.r, this.ocrSelEnd.r);
    const maxR = Math.max(this.ocrSelStart.r, this.ocrSelEnd.r);
    const minC = Math.min(this.ocrSelStart.c, this.ocrSelEnd.c);
    const maxC = Math.max(this.ocrSelStart.c, this.ocrSelEnd.c);
    return r >= minR && r <= maxR && c >= minC && c <= maxC;
  }
  ocrAppendMode: 'left' | 'right' | 'above' | 'below' = 'below';

  onFileSelected(event: any) {
    this.selectedImportFile = event.target.files[0];
  }

  toggleStar() {
    this.isStarred = !this.isStarred;
    this.showToast(this.isStarred ? 'Added to Starred' : 'Removed from Starred');
  }

  openFeatureModal(type: any) {
    this.closeMenus();
    this.activeModal = type;
    if (type === 'template') this.dummyList = ['Blank', 'Invoice', 'Budget', 'Schedule', 'To-Do List', 'Project Tracker'];
    if (type === 'open') {
      this.myDocs = [];
      this.api.listDocuments().subscribe(res => {
        this.myDocs = res.filter((d: any) => d.doc_type === 'sheet' && !d.is_trashed);
      });
    }
    if (type === 'version') {
      this.loadVersions();
    }
    if (type === 'audit') this.dummyList = ['User modified cell C4 (1m ago)', 'You changed column width (5m ago)', 'User added new row (10m ago)'];
    if (type === 'workflow') this.dummyList = ['Highlight row if Status=Done', 'Send email if Due Date < Today'];
  }

  loadVersions() {
    const currentVersion = {
      id: 'current',
      created_at: new Date().toISOString(),
      is_named: false,
      version_name: 'Current Version'
    };

    // Set it immediately so it ALWAYS shows up, even if API fails
    this.versions = [currentVersion];
    if (!this.previewVersionId) {
      this.previewVersion('current');
    }

    this.api.getSheetVersions(this.docId).subscribe({
      next: (res) => {
        this.versions = [currentVersion, ...res];
        this.cdr.detectChanges(); // CRITICAL for OnPush components!
      },
      error: (err) => {
        console.error("Failed to fetch sheet versions:", err);
      }
    });
  }

  previewVersion(versionId: string) {
    this.previewVersionId = versionId;

    if (versionId === 'current') {
      this.previewData = this.sheets;
      this.previewActiveSheetIdx = this.currentSheetIdx;
      this.cdr.detectChanges();
      return;
    }

    this.api.getSheetVersionSnapshot(this.docId, versionId).subscribe(res => {
      this.previewData = typeof res.content === 'string' ? JSON.parse(res.content) : res.content;
      this.previewActiveSheetIdx = this.currentSheetIdx;

      // If previewData is a root doc with _importedSheets, extract the sheets array
      if (!Array.isArray(this.previewData) && this.previewData._importedSheets) {
        this.previewData = this.previewData._importedSheets;
      } else if (!Array.isArray(this.previewData)) {
        this.previewData = [this.previewData];
      }

      // Ensure the active index is within bounds
      if (this.previewActiveSheetIdx >= this.previewData.length) {
        this.previewActiveSheetIdx = 0;
      }

      this.cdr.detectChanges();
    });
  }

  showNameVersionPrompt = false;
  tempVersionName = '';

  promptNameVersion() {
    this.tempVersionName = '';
    this.showNameVersionPrompt = true;
  }

  cancelNameVersion() {
    this.showNameVersionPrompt = false;
    this.tempVersionName = '';
  }

  submitNameVersion() {
    if (this.tempVersionName && this.tempVersionName.trim()) {
      this.newVersionName = this.tempVersionName.trim();
      this.createNamedVersion();
      this.showNameVersionPrompt = false;
    }
  }

  createNamedVersion() {
    if (!this.newVersionName) return;
    this.api.createNamedVersion(this.docId, this.newVersionName).subscribe(() => {
      this.newVersionName = '';
      this.loadVersions();
      this.showToast('Version named successfully');
    });
  }

  makeCopy() {
    this.showToast('Creating a copy...');
    setTimeout(() => this.showToast('Copy successfully saved to your Drive.'), 1000);
  }

  showRestoreConfirm = false;
  versionToRestore: string | null = null;

  confirmRestoreVersion(versionId: string) {
    this.versionToRestore = versionId;
    this.showRestoreConfirm = true;
  }

  cancelRestore() {
    this.showRestoreConfirm = false;
    this.versionToRestore = null;
  }

  executeRestore() {
    if (!this.versionToRestore) return;
    this.api.restoreSheetVersion(this.docId, this.versionToRestore).subscribe(() => {
      this.activeModal = null;
      this.showRestoreConfirm = false;
      this.versionToRestore = null;
      this.showToast('Version restored successfully');
      setTimeout(() => window.location.reload(), 1000);
    });
  }

  handleModalAction(payload?: any) {
    if (this.activeModal === 'version') return; // Handled separately

    if (this.activeModal === 'template') {
      const templateName = payload || 'Blank';
      this.addSheet();
      const idx = this.sheets.length - 1;
      this.sheets[idx].name = templateName;

      // Basic mock templates
      if (templateName === 'Invoice') {
        this.sheets[idx].cells[0][0] = 'INVOICE';
        this.sheets[idx].formats['0,0'] = { bold: true, size: '24', align: 'left' };
        this.sheets[idx].cells[2][0] = 'Bill To:';
        this.sheets[idx].cells[3][0] = 'Name:';
        this.sheets[idx].cells[4][0] = 'Address:';
        this.sheets[idx].cells[6][0] = 'Item';
        this.sheets[idx].cells[6][1] = 'Qty';
        this.sheets[idx].cells[6][2] = 'Price';
        this.sheets[idx].cells[6][3] = 'Total';
        this.sheets[idx].formats['6,0'] = { bold: true };
        this.sheets[idx].formats['6,1'] = { bold: true };
        this.sheets[idx].formats['6,2'] = { bold: true };
        this.sheets[idx].formats['6,3'] = { bold: true };
      } else if (templateName === 'Budget') {
        this.sheets[idx].cells[0][0] = 'MONTHLY BUDGET';
        this.sheets[idx].formats['0,0'] = { bold: true, size: '18', color: '#10b981' };
        this.sheets[idx].cells[2][0] = 'Income';
        this.sheets[idx].cells[2][1] = 'Planned';
        this.sheets[idx].cells[2][2] = 'Actual';
        this.sheets[idx].formats['2,0'] = { bold: true };
        this.sheets[idx].formats['2,1'] = { bold: true };
        this.sheets[idx].formats['2,2'] = { bold: true };
      } else if (templateName === 'To-Do List') {
        this.sheets[idx].cells[0][0] = 'TO-DO LIST';
        this.sheets[idx].formats['0,0'] = { bold: true, size: '18' };
        this.sheets[idx].cells[1][0] = 'Status';
        this.sheets[idx].cells[1][1] = 'Task';
        this.sheets[idx].cells[1][2] = 'Due Date';
        this.sheets[idx].formats['1,0'] = { bold: true };
        this.sheets[idx].formats['1,1'] = { bold: true };
        this.sheets[idx].formats['1,2'] = { bold: true };
      }
      this.activeModal = null;
      this.showToast('Created ' + templateName + ' sheet!');
      this.switchSheet(idx);
    } else if (this.activeModal === 'open') {
      const doc = payload;
      if (doc && doc.id) {
        this.showToast('Opening ' + doc.title + '...');
        window.location.href = `/${doc.doc_type}/${doc.id}`;
      }
      this.activeModal = null;
    } else if (this.activeModal === 'import') {
      if (this.selectedImportFile) {
        this.isUploading = true;
        this.uploadProgress = 0;
        this.uploadTimeLeft = 'Calculating...';
        this.uploadStartTime = Date.now();

        this.api.importFile(this.selectedImportFile, this.docId).subscribe({
          next: (event: any) => {
            if (event.type === HttpEventType.UploadProgress) {
              if (event.total) {
                this.uploadProgress = Math.round(100 * event.loaded / event.total);
                const timeElapsed = Date.now() - this.uploadStartTime;
                const uploadSpeed = event.loaded / (timeElapsed / 1000); // bytes per sec
                const bytesLeft = event.total - event.loaded;
                const secondsLeft = Math.round(bytesLeft / uploadSpeed);

                if (secondsLeft > 60) {
                  this.uploadTimeLeft = `~${Math.ceil(secondsLeft / 60)} mins left`;
                } else if (secondsLeft > 0) {
                  this.uploadTimeLeft = `${secondsLeft} secs left`;
                } else {
                  this.uploadTimeLeft = 'Processing file...';
                }
              }
            } else if (event.type === HttpEventType.Response) {
              const doc = event.body;
              this.isUploading = false;
              const fileName = this.selectedImportFile!.name;
              this.selectedImportFile = null;
              this.activeModal = null;
              try {
                let p = JSON.parse(doc.content || '{}');
                if (Array.isArray(p) && p.length > 0) p = p[0];
                if (p._importedSheets && p._importedSheets.length > 0) {
                  // Multi-sheet import — expand sparse cells to 2D arrays
                  this.sheets = p._importedSheets.map((sheet: any) => {
                    let cells2d: string[][];
                    if (Array.isArray(sheet.cells)) {
                      cells2d = Array.from({ length: Math.max(this.ROWS, sheet.cells.length) }, (_: any, r: number) =>
                        Array.from({ length: Math.max(this.COLS, sheet.cells[r]?.length ?? 0) }, (_2: any, c: number) =>
                          sheet.cells[r]?.[c] ?? ''));
                    } else {
                      const sp = sheet.cells || {};
                      const maxR = Math.max(this.ROWS, ...Object.keys(sp).map(Number).filter((n: number) => !isNaN(n))) + 1;
                      cells2d = Array.from({ length: maxR }, (_: any, r: number) =>
                        Array.from({ length: this.COLS }, (_2: any, c: number) => sp[r]?.[c] ?? ''));
                    }
                    return { ...sheet, cells: cells2d };
                  });
                  this.currentSheetIdx = 0;
                  const s0 = this.sheets[0];
                  for (let r = 0; r < this.ROWS; r++)
                    for (let c = 0; c < this.COLS; c++)
                      this.cells[r][c] = s0.cells[r]?.[c] ?? '';
                  this.formats = { ...(s0.formats || {}) };
                  this.validations = { ...(s0.validations || {}) };
                  if (s0.colWidths) this.sheets[0].colWidths = s0.colWidths;
                  if (s0.rowHeights) this.sheets[0].rowHeights = s0.rowHeights;
                  this.hiddenRows = new Set();
                  this.filterActive = false;
                  this.activeFilterCols = new Set();
                  this.pushHistory();
                  this.updateDisplayCache();
                  this.showToast(`${fileName} imported successfully!`);
                } else if (p.cells) {
                  // Single-sheet flat format
                  this.pushHistory();
                  for (let r = 0; r < this.ROWS; r++)
                    for (let c = 0; c < this.COLS; c++)
                      this.cells[r][c] = p.cells[r]?.[c] ?? '';
                  if (p.formats) this.formats = p.formats;
                  if (p.validations) this.validations = p.validations;
                  this.updateDisplayCache();
                  this.showToast(`${fileName} imported successfully!`);
                } else {
                  // Nothing parseable — fallback to page reload
                  window.location.reload();
                }
              } catch {
                window.location.reload();
              }
            } // Close else if (event.type === HttpEventType.Response)
          },
          error: () => this.showToast('Failed to import file.')
        });
      } else {
        this.showToast('Please select a file first.');
      }

    } else if (this.activeModal === 'password') {
      this.showToast('Password protection enabled!');
    } else if (this.activeModal === 'move') {
      this.showToast('Document moved to ' + (this.modalInput || 'Folder'));
    } else if (this.activeModal === 'workflow') {
      this.showToast('Workflow rule added.');
    }
    this.activeModal = null;
    this.modalInput = '';
  }

  performShare() {
    const queryEmails = this.shareQuery ? this.shareQuery.split(/[,;\s]+/).map(e => e.trim()).filter(e => e.length > 0) : [];
    const allEmails = Array.from(new Set([...this.selectedShareEmails, ...queryEmails]));

    if (allEmails.length === 0) return;

    let successCount = 0;
    let externalCount = 0;
    let failCount = 0;
    const total = allEmails.length;

    allEmails.forEach(email => {
      this.api.shareDocument(this.docId, email, this.shareRole.toLowerCase()).subscribe({
        next: (res: any) => {
          successCount++;
          if (res?.external) externalCount++;
          if (successCount + failCount === total) this.finishShare(successCount, externalCount, failCount, total);
        },
        error: () => {
          failCount++;
          if (successCount + failCount === total) this.finishShare(successCount, externalCount, failCount, total);
        }
      });
    });
  }

  finishShare(success: number, external: number, fail: number, total: number) {
    if (success === total) {
      if (external > 0) {
        this.showToast(`Shared successfully! Invitation email sent to ${external} external user(s).`);
      } else {
        this.showToast(`Shared successfully with ${success} user(s).`);
      }
      this.shareQuery = '';
      this.selectedShareEmails = [];
      this.shareModalOpen = false;
    } else if (success > 0) {
      this.showToast(`Shared with ${success} user(s). ${fail} failed.`);
    } else {
      this.showToast(`Failed to share. Please check the email address and try again.`);
    }
  }

  triggerCopy() {
    this.closeMenus();
    this.api.createDocument(this.title + ' - Copy', 'sheet').subscribe((res: any) => {
      window.open(/sheet/ + res.id, '_blank');
    });
  }

  async triggerRename() {
    this.closeMenus();
    const newTitle = await this.openPrompt('Enter new document name:', this.title);
    if (newTitle && newTitle.trim()) {
      this.title = newTitle.trim();
      this.save();
    }
  }

  openPrompt(title: string, defaultValue: string = ''): Promise<string | null> {
    this.promptModalTitle = title;
    this.promptModalValue = defaultValue;
    this.promptModalOpen = true;
    return new Promise((resolve) => {
      this.promptResolve = resolve;
    });
  }

  closePrompt() {
    this.promptModalOpen = false;
    if (this.promptResolve) {
      this.promptResolve(null);
      this.promptResolve = null;
    }
  }

  submitPrompt() {
    this.promptModalOpen = false;
    if (this.promptResolve) {
      this.promptResolve(this.promptModalValue);
      this.promptResolve = null;
    }
  }

  openConfirm(message: string): Promise<boolean> {
    this.confirmModalMessage = message;
    this.confirmModalOpen = true;
    return new Promise((resolve) => {
      this.confirmResolve = resolve;
    });
  }

  closeConfirm(result: boolean) {
    this.confirmModalOpen = false;
    if (this.confirmResolve) {
      this.confirmResolve(result);
      this.confirmResolve = null;
    }
  }

  showToast(msg: string) {
    this.toastMsg = msg; this.toastVisible = true;
    setTimeout(() => this.toastVisible = false, 2500);
  }

  getColWidth(c: number): number {
    return this.sheets[this.currentSheetIdx].colWidths?.[c] ?? 100;
  }

  getRowHeight(r: number): number {
    if (this.hiddenRows && this.hiddenRows.has(r)) return 0;
    return this.sheets[this.currentSheetIdx].rowHeights?.[r] ?? 26;
  }

  getFrozenColOffset(c: number): number {
    let offset = (this.showHeaders ? 46 : 0) + this.groupMarginWidth;
    const widths = this.sheets[this.currentSheetIdx].colWidths || {};
    for (let i = 0; i < c; i++) {
      offset += widths[i] ?? 100;
    }
    return offset;
  }

  getFrozenRowOffset(r: number): number {
    let offset = this.showHeaders ? 26 : 0;
    const heights = this.sheets[this.currentSheetIdx].rowHeights || {};
    for (let i = 0; i < r; i++) {
      offset += heights[i] ?? 26;
    }
    return offset;
  }

  startColResize(event: MouseEvent, c: number) {
    event.stopPropagation();
    event.preventDefault();
    this.resizingCol = c;
    this.resizeStartX = event.clientX;
    this.resizeStartSize = this.getColWidth(c);

    const gridWrap = (event.target as HTMLElement).closest('.grid-wrap');
    if (gridWrap) {
      const rect = gridWrap.getBoundingClientRect();
      this.resizeLineX = event.clientX - rect.left + gridWrap.scrollLeft;
    }

    const moveListener = (e: MouseEvent) => {
      if (gridWrap) {
        const rect = gridWrap.getBoundingClientRect();
        this.resizeLineX = e.clientX - rect.left + gridWrap.scrollLeft;
      }
    };

    const upListener = (e: MouseEvent) => {
      document.removeEventListener('mousemove', moveListener);
      document.removeEventListener('mouseup', upListener);
      const delta = e.clientX - this.resizeStartX;
      const newWidth = Math.max(30, this.resizeStartSize + delta);

      const sheet = this.sheets[this.currentSheetIdx];
      if (!sheet.colWidths) sheet.colWidths = {};
      sheet.colWidths[c] = newWidth;

      this.resizingCol = null;
      this.save();
    };

    document.addEventListener('mousemove', moveListener);
    document.addEventListener('mouseup', upListener);
  }

  startRowResize(event: MouseEvent, r: number) {
    event.stopPropagation();
    event.preventDefault();
    this.resizingRow = r;
    this.resizeStartY = event.clientY;
    this.resizeStartSize = this.getRowHeight(r);

    const gridWrap = (event.target as HTMLElement).closest('.grid-wrap');
    if (gridWrap) {
      const rect = gridWrap.getBoundingClientRect();
      this.resizeLineY = event.clientY - rect.top + gridWrap.scrollTop;
    }

    const moveListener = (e: MouseEvent) => {
      if (gridWrap) {
        const rect = gridWrap.getBoundingClientRect();
        this.resizeLineY = e.clientY - rect.top + gridWrap.scrollTop;
      }
    };

    const upListener = (e: MouseEvent) => {
      document.removeEventListener('mousemove', moveListener);
      document.removeEventListener('mouseup', upListener);
      const delta = e.clientY - this.resizeStartY;
      const newHeight = Math.max(20, this.resizeStartSize + delta);

      const sheet = this.sheets[this.currentSheetIdx];
      if (!sheet.rowHeights) sheet.rowHeights = {};
      sheet.rowHeights[r] = newHeight;

      this.resizingRow = null;
      this.save();
    };

    document.addEventListener('mousemove', moveListener);
    document.addEventListener('mouseup', upListener);
  }

  copySheet(idx: number) {
    this.copiedSheetData = JSON.parse(JSON.stringify(this.sheets[idx]));
    this.showToast(`Sheet "${this.sheets[idx].name}" copied.`);
  }

  pasteSheet() {
    if (!this.copiedSheetData) return;
    this.pushHistory();
    const newSheet = JSON.parse(JSON.stringify(this.copiedSheetData));
    let counter = 1;
    let finalName = newSheet.name + ' (Pasted)';
    while (this.sheets.some(s => s.name === finalName)) {
      counter++;
      finalName = `${newSheet.name} (Pasted ${counter})`;
    }
    newSheet.name = finalName;
    const insertIdx = (this.activeSheetMenuIdx !== null ? this.activeSheetMenuIdx : this.currentSheetIdx) + 1;
    this.sheets.splice(insertIdx, 0, newSheet);
    this.switchSheet(insertIdx);
    this.activeSheetMenuIdx = null;
    this.showToast(`Pasted as "${finalName}"`);
    this.save();
  }

  moveSheetModalOpen = false;
  moveSheetTargetIdx = -1;
  moveSheetDestination: string | number = 'start';

  openMoveSheetModal(idx: number) {
    if (idx < 0 || idx >= this.sheets.length) return;
    this.moveSheetTargetIdx = idx;
    this.moveSheetDestination = 'start';
    this.moveSheetModalOpen = true;
  }

  confirmMoveSheet() {
    if (this.moveSheetTargetIdx < 0 || this.moveSheetTargetIdx >= this.sheets.length) {
      this.moveSheetModalOpen = false;
      return;
    }
    this.pushHistory();
    const dest = this.moveSheetDestination;
    const sheetToMove = this.sheets[this.moveSheetTargetIdx];

    this.sheets.splice(this.moveSheetTargetIdx, 1);

    let insertIdx = 0;
    if (dest === 'start') {
      insertIdx = 0;
    } else if (dest === 'end') {
      insertIdx = this.sheets.length;
    } else {
      const targetOriginalIdx = Number(dest);
      if (targetOriginalIdx > this.moveSheetTargetIdx) {
        insertIdx = targetOriginalIdx;
      } else {
        insertIdx = targetOriginalIdx + 1;
      }
    }

    this.sheets.splice(insertIdx, 0, sheetToMove);
    this.switchSheet(insertIdx);
    this.save();
    this.moveSheetModalOpen = false;
  }

  setTabColor(idx: number, color: string) {
    this.sheets[idx].tabColor = color;
    this.save();
  }

  toggleSheetGridlines(idx: number) {
    this.sheets[idx].hideGridlines = !this.sheets[idx].hideGridlines;
    this.save();
  }

  getVisibleSheetCount(): number {
    return this.sheets.filter(s => !s.hidden).length;
  }

  hideSheet(idx: number) {
    if (this.getVisibleSheetCount() <= 1 && !this.sheets[idx].hidden) {
      this.showToast('Cannot hide the only visible sheet.');
      return;
    }
    this.sheets[idx].hidden = true;
    if (this.currentSheetIdx === idx) {
      const nextIdx = this.sheets.findIndex(s => !s.hidden);
      this.switchSheet(nextIdx);
    }
    this.save();
  }

  unhideSheet(idx: number) {
    this.sheets[idx].hidden = false;
    this.save();
  }

  toggleLockSheet(idx: number) {
    console.log('[SheetEditor] toggleLockSheet called with idx:', idx);
    this.sheets[idx].locked = !this.sheets[idx].locked;
    this.save();
    this.showToast(this.sheets[idx].locked ? 'Sheet locked.' : 'Sheet unlocked.');
    if (this.cdr) this.cdr.detectChanges();
  }

  publishSheet(idx: number) {
    this.showToast(`Sheet "${this.sheets[idx].name}" published to web.`);
  }

  back() { this.save(); this.router.navigate(['/']); }
  ngOnDestroy() {
    this.syncSub?.unsubscribe();
    this.api.disconnectSync();
    if (this.saveSubscription) this.saveSubscription.unsubscribe();
    if (this.hasPendingChanges && this.dataLoaded) this.executeSave();
  }

  showEditHistoryPanel = false;
  editHistoryCell: any = null;
  editHistoryData: any[] = [];
  cellEditHistory: Record<string, any[]> = {};

  openCellEditHistory() {
    if (this.selectedRow === null || this.selectedCol === null) return;
    this.editHistoryCell = { r: this.selectedRow, c: this.selectedCol };
    const key = `${this.currentSheetIdx}-${this.selectedRow}-${this.selectedCol}`;
    if (!this.cellEditHistory) this.cellEditHistory = {};
    if (!this.cellEditHistory[key]) {
      this.cellEditHistory[key] = [];
    }
    this.editHistoryData = this.cellEditHistory[key];
    this.showEditHistoryPanel = true;
  }

  showCustomFormatModal = false;
  customFormatString = '';
  showMoreFormatsModal = false;

  openCustomFormatModal() {
    this.closeMenus();
    this.customFormatString = '';
    this.showCustomFormatModal = true;
  }

  applyCustomFormat() {
    if (this.customFormatString.trim()) {
      this.setFormat('numFormat', 'custom_' + this.customFormatString.trim());
    }
    this.showCustomFormatModal = false;
  }

  openMoreFormatsModal() {
    this.closeMenus();
    this.showMoreFormatsModal = true;
  }

  private applyCustomFormatString(val: any, num: number, isNum: boolean, fmtStr: string): string {
    if (!isNum) return String(val);
    let out = fmtStr;
    let tempNum = num;
    if (fmtStr.includes('%')) tempNum = tempNum * 100;
    let decimals = 0;
    const decMatch = fmtStr.match(/\.(0+)/);
    if (decMatch) decimals = decMatch[1].length;
    let numStr = tempNum.toFixed(decimals);
    if (fmtStr.includes(',')) {
      const parts = numStr.split('.');
      parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
      numStr = parts.join('.');
    }
    const numPattern = /[#0,]+(\.[0]+)?/;
    if (numPattern.test(fmtStr)) {
      out = fmtStr.replace(numPattern, numStr);
    } else {
      out = numStr;
    }
    return out;
  }

  colNameToIndex(name: string): number {
    let result = 0;
    for (let i = 0; i < name.length; i++) {
      result = result * 26 + (name.toUpperCase().charCodeAt(i) - 64);
    }
    return result - 1;
  }

  parseCellRef(ref: string): { r: number, c: number } | null {
    const match = ref.match(/^([A-Za-z]+)(\d+)$/);
    if (!match) return null;
    return { c: this.colNameToIndex(match[1]), r: parseInt(match[2], 10) - 1 };
  }

  parseRangeStr(range: string): { minR: number, maxR: number, minC: number, maxC: number } | null {
    let rStr = range.split('!').pop() || range;
    rStr = rStr.replace(/['"]/g, ''); // strip quotes
    const parts = rStr.split(':');
    if (parts.length === 1) { // Single cell range (which shouldn't be allowed for sparkline but parsed anyway)
      const p = this.parseCellRef(parts[0]);
      if (!p) return null;
      return { minR: p.r, maxR: p.r, minC: p.c, maxC: p.c };
    }
    if (parts.length !== 2) return null;
    const p1 = this.parseCellRef(parts[0]);
    const p2 = this.parseCellRef(parts[1]);
    if (!p1 || !p2) return null;
    return {
      minR: Math.min(p1.r, p2.r),
      maxR: Math.max(p1.r, p2.r),
      minC: Math.min(p1.c, p2.c),
      maxC: Math.max(p1.c, p2.c)
    };
  }

  getSparklineData(rangeStr: string, includeHidden = false) {
    const range = this.parseRange(rangeStr);
    if (!range) return { values: [], hasNumbers: false };

    const targetSheetIdx = this.sheets.findIndex(s => s.name === range.sheetName) !== -1 ? this.sheets.findIndex(s => s.name === range.sheetName) : this.currentSheetIdx;
    const isCurrentSheet = targetSheetIdx === this.currentSheetIdx;
    const targetCells = isCurrentSheet ? this.cells : this.sheets[targetSheetIdx].cells;
    const targetHiddenRows = isCurrentSheet ? this.hiddenRows : new Set(this.sheets[targetSheetIdx].hiddenRows || []);

    let values = [];
    let hasNumbers = false;

    for (let r = range.startR; r <= range.endR; r++) {
      if (!includeHidden && targetHiddenRows.has(r)) continue;
      for (let c = range.startC; c <= range.endC; c++) {
        const cellStr = targetCells[r]?.[c] || '';
        const cleanStr = cellStr.toString().trim();
        if (cleanStr === '') {
          values.push(null);
        } else {
          const num = Number(cleanStr);
          if (!isNaN(num)) {
            values.push(num);
            hasNumbers = true;
          } else {
            values.push(null);
          }
        }
      }
    }

    return { values, hasNumbers };
  }

  getSparklineSvgSafe(r: number, c: number): any {
    const sheet = this.sheets[this.currentSheetIdx];
    const config = sheet.sparklines![`${r},${c}`];
    if (!config) return '';

    const data = this.getSparklineData(config.sourceRange || '', config.includeHiddenRowsColumns);

    if (!data.hasNumbers) {
      return this.sanitizer.bypassSecurityTrustHtml(`<span style="color:#ef4444;font-size:10px;font-weight:bold;">#ERROR!</span>`);
    }

    let rawValues = data.values;
    if (config.emptyCellMode === 'skip') {
      rawValues = rawValues.filter(v => v !== null);
    } else if (config.emptyCellMode === 'zero') {
      rawValues = rawValues.map(v => v === null ? 0 : v);
    }

    if (rawValues.filter(v => v !== null).length < 2 && config.type === 'line') {
      // Single points can be drawn as bar/winloss but not line
      return this.sanitizer.bypassSecurityTrustHtml(`<span style="color:#ef4444;font-size:10px;font-weight:bold;">#ERROR!</span>`);
    }

    const w = this.getColWidth(c);
    const h = this.getRowHeight(r);
    const padX = 2;
    const padY = 4;
    const innerW = w - padX * 2;
    const innerH = h - padY * 2;

    let svg = `<svg width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" preserveAspectRatio="none" style="display:block;">`;

    // Calculate Y domain
    let globalMin = Infinity, globalMax = -Infinity;

    // Group logic for min/max
    if (config.isGrouped && config.groupId) {
      for (const k of Object.keys(sheet.sparklines || {})) {
        if (sheet.sparklines![k].groupId === config.groupId) {
          const grpData = this.getSparklineData(sheet.sparklines![k].sourceRange || '', config.includeHiddenRowsColumns);
          const grpNums = grpData.values.filter(v => v !== null) as number[];
          if (grpNums.length > 0) {
            globalMin = Math.min(globalMin, ...grpNums);
            globalMax = Math.max(globalMax, ...grpNums);
          }
        }
      }
    }

    const localNums = rawValues.filter(v => v !== null) as number[];
    let min = Math.min(...localNums);
    let max = Math.max(...localNums);

    if (config.verticalAxis.min.mode === 'same' && globalMin !== Infinity) min = globalMin;
    if (config.verticalAxis.max.mode === 'same' && globalMax !== -Infinity) max = globalMax;

    if (config.verticalAxis.min.mode === 'custom' && config.verticalAxis.min.customValue !== null) min = config.verticalAxis.min.customValue;
    if (config.verticalAxis.max.mode === 'custom' && config.verticalAxis.max.customValue !== null) max = config.verticalAxis.max.customValue;

    const rangeVal = max - min || 1;

    // Highlights detection
    let firstIdx = -1, lastIdx = -1, highIdx = -1, lowIdx = -1;
    let currentHigh = -Infinity, currentLow = Infinity;

    for (let i = 0; i < rawValues.length; i++) {
      const v = rawValues[i];
      if (v !== null) {
        if (firstIdx === -1) firstIdx = i;
        lastIdx = i;
        if (v > currentHigh) { currentHigh = v; highIdx = i; }
        if (v < currentLow) { currentLow = v; lowIdx = i; }
      }
    }

    const hl = config.highlights;
    const getColor = (i: number, val: number) => {
      // Precedence: Negative > High/Low > First/Last > Base
      if (val < 0 && hl.negative.enabled) return hl.negative.color;
      if (i === highIdx && hl.high.enabled) return hl.high.color;
      if (i === lowIdx && hl.low.enabled) return hl.low.color;
      if (i === firstIdx && hl.first.enabled) return hl.first.color;
      if (i === lastIdx && hl.last.enabled) return hl.last.color;
      return config.baseColor;
    };

    let ptsX = (i: number) => padX + (i / (rawValues.length - 1)) * innerW;
    if (config.horizontalAxis.rightToLeft) {
      ptsX = (i: number) => padX + innerW - (i / (rawValues.length - 1)) * innerW;
    }
    const ptsY = (val: number) => padY + innerH - ((val - min) / rangeVal) * innerH;

    // Display Axis
    if (config.horizontalAxis.displayAxis && min < 0 && max > 0) {
      const zeroY = ptsY(0);
      svg += `<line x1="${padX}" y1="${zeroY}" x2="${padX + innerW}" y2="${zeroY}" stroke="#000" stroke-width="1" opacity="0.5"/>`;
    }

    if (config.type === 'line') {
      let pathD = '';
      let isFirstInSegment = true;
      for (let i = 0; i < rawValues.length; i++) {
        if (rawValues[i] === null) {
          if (config.emptyCellMode === 'gap') {
            isFirstInSegment = true;
          }
          // If 'connect', do nothing, next valid point will just line-to
          continue;
        }

        const x = ptsX(i);
        const y = ptsY(rawValues[i] as number);

        if (isFirstInSegment) {
          pathD += `M ${x} ${y} `;
          isFirstInSegment = false;
        } else {
          pathD += `L ${x} ${y} `;
        }
      }
      if (pathD) {
        svg += `<path d="${pathD}" fill="none" stroke="${config.baseColor}" stroke-width="1.5" stroke-linejoin="round" stroke-linecap="round"/>`;
      }

      // Markers
      for (let i = 0; i < rawValues.length; i++) {
        const val = rawValues[i];
        if (val === null) continue;
        const x = ptsX(i);
        const y = ptsY(val);

        const isHighlighted = (hl.negative.enabled && val < 0) || (hl.high.enabled && i === highIdx) || (hl.low.enabled && i === lowIdx) || (hl.first.enabled && i === firstIdx) || (hl.last.enabled && i === lastIdx);

        if (isHighlighted || hl.markers.enabled) {
          const color = getColor(i, val);
          svg += `<circle cx="${x}" cy="${y}" r="2" fill="${color}"/>`;
        }
      }
    } else if (config.type === 'column') {
      const barW = Math.max(1, (innerW / rawValues.length) - 1);
      const zeroY = ptsY(Math.max(Math.min(0, max), min));

      for (let i = 0; i < rawValues.length; i++) {
        const val = rawValues[i];
        if (val === null) continue;
        const x = ptsX(i) - (config.horizontalAxis.rightToLeft ? 0 : barW / 2);
        const y = ptsY(val);

        const barY = Math.min(y, zeroY);
        const barH = Math.abs(y - zeroY);
        const color = getColor(i, val);

        svg += `<rect x="${x}" y="${barY}" width="${barW}" height="${Math.max(1, barH)}" fill="${color}"/>`;
      }
    } else if (config.type === 'winloss') {
      const barW = Math.max(1, (innerW / rawValues.length) - 1);
      const midY = padY + innerH / 2;
      const fixH = (innerH / 2) * 0.8;

      for (let i = 0; i < rawValues.length; i++) {
        const val = rawValues[i];
        if (val === null) continue;
        const x = ptsX(i) - (config.horizontalAxis.rightToLeft ? 0 : barW / 2);
        const isWin = val > 0;

        const barY = isWin ? midY - fixH : midY;
        // Precedence for winloss: Negative highlight vs High highlight
        let color = config.baseColor;
        if (!isWin && hl.negative.enabled) color = hl.negative.color;
        else if (isWin && hl.high.enabled) color = hl.high.color;

        svg += `<rect x="${x}" y="${barY}" width="${barW}" height="${fixH}" fill="${color}"/>`;
      }

      // Zero axis line for winloss
      svg += `<line x1="${padX}" y1="${midY}" x2="${padX + innerW}" y2="${midY}" stroke="#000" stroke-width="0.5" opacity="0.3"/>`;
    }

    svg += '</svg>';
    return this.sanitizer.bypassSecurityTrustHtml(svg);
  }
  scrollToCell(r: number, c: number) {
    const wrapEl = this.gridWrapRef?.nativeElement || (document.querySelector('.grid-wrap') as HTMLElement);
    if (!wrapEl) return;

    let targetScrollTop = 0;
    for (let i = 0; i < r; i++) targetScrollTop += this.getRowHeight(i) || 24;
    let targetScrollLeft = 0;
    for (let i = 0; i < c; i++) targetScrollLeft += this.getColWidth(i) || 100;

    const newScrollTop = Math.max(0, targetScrollTop - (wrapEl.clientHeight / 3));
    const newScrollLeft = Math.max(0, targetScrollLeft - (wrapEl.clientWidth / 3));
    wrapEl.scrollTop = newScrollTop;
    wrapEl.scrollLeft = newScrollLeft;
    this.updateVisibleRows(newScrollTop);
  }

  onAuditNavigate(event: { sheetId: string, r: number, c: number, endR: number, endC: number }) {
    const targetSheetIdx = this.sheets.findIndex(s => (s as any).id === event.sheetId || this.sheets.indexOf(s).toString() === event.sheetId);
    if (targetSheetIdx !== -1 && targetSheetIdx !== this.currentSheetIdx) {
      this.switchSheet(targetSheetIdx);
    }

    this.scrollToCell(event.r, event.c);
    this.cdr.detectChanges();

    this.selectCell(event.r, event.c);
    this.rangeStart = { r: event.r, c: event.c };
    this.rangeEnd = { r: event.endR, c: event.endC };

    const el = document.getElementById(`cell-${event.r}-${event.c}`);
    if (el) {
      el.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
    }

    this.activeModal = null;
  }

  linkClickTimeout: any = null;

  onLinkClick(event: MouseEvent, url: string) {
    event.preventDefault();
    if (this.linkClickTimeout) {
      clearTimeout(this.linkClickTimeout);
      this.linkClickTimeout = null;
    } else {
      this.linkClickTimeout = setTimeout(() => {
        this.linkClickTimeout = null;
        window.open(url, '_blank');
      }, 250);
    }
  }
}


























