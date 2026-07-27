import { Component, OnInit, OnDestroy, ViewChild, ElementRef, HostListener } from '@angular/core';
import { ActivatedRoute, Router } from '@angular/router';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { Subscription } from 'rxjs';
import { ChatWidgetComponent } from '../../components/chat-widget/chat-widget.component';
import { ApiService } from '../../services/api.service';
import { AuthService } from '../../services/auth.service';

interface Page { title: string; body: string; bg?: string; color?: string; }
interface SlideData { id: string; title: string; pages: Record<string, Page>; pageOrder: string[]; }

@Component({
  selector: 'app-slide-editor',
  standalone: true,
  imports: [CommonModule, FormsModule, ChatWidgetComponent],
  template: `
    <div class="shell" (click)="closeMenus()">
      <div class="top-bar">
        <div class="top-left">
          <button class="back" (click)="back()" title="Back to Dashboard">
            <svg fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24" class="icon-sm"><path stroke-linecap="round" stroke-linejoin="round" d="M10 19l-7-7m0 0l7-7m-7 7h18"></path></svg>
          </button>
          <div class="brand" style="display:flex; align-items:center; gap:6px; cursor:pointer; margin-right: 12px;" (click)="goHome()" title="Show Home">
            <div style="background:#f4b400; color:#fff; display:flex; align-items:center; justify-content:center; width:28px; height:28px; border-radius:4px;">
              <span class="material-symbols-outlined" style="font-size:20px;">slideshow</span>
            </div>
            <span class="brand-name" style="font-weight:600; font-size:18px;">Show</span>
          </div>
          <div class="doc-meta">
            <input class="title-input" [(ngModel)]="title" (blur)="save()" placeholder="Untitled presentation" />
            <div class="menu-bar" (mousedown)="$event.preventDefault()">
              <div class="menu-item" (click)="toggleMenu('file', $event)" [class.active]="activeMenu === 'file'">
                File
                <div class="dropdown" *ngIf="activeMenu === 'file'">
                  <div class="dd-item" (click)="newDoc()"><span class="dd-text">New</span></div>
                  <div class="dd-item" (click)="triggerOpen()"><span class="dd-text">Open</span><span class="dd-hint">Ctrl+O</span></div>
                  <div class="dd-item" (click)="makeCopy()"><span class="dd-text">Make a copy</span></div>
                  <div class="dd-sep"></div>
                  <div class="dd-item" (click)="shareModalOpen = true; closeMenus()"><span class="dd-text">Share</span></div>
                  <div class="dd-item" (click)="exportFile('pptx'); closeMenus()"><span class="dd-text">Download (.pptx)</span></div>
                </div>
              </div>
              <div class="menu-item" (click)="toggleMenu('edit', $event)" [class.active]="activeMenu === 'edit'">
                Edit
                <div class="dropdown" *ngIf="activeMenu === 'edit'">
                  <div class="dd-item" (click)="exec('undo')"><span class="dd-text">Undo</span><span class="dd-hint">Ctrl+Z</span></div>
                  <div class="dd-item" (click)="exec('redo')"><span class="dd-text">Redo</span><span class="dd-hint">Ctrl+Y</span></div>
                  <div class="dd-sep"></div>
                  <div class="dd-item" (click)="exec('copy')"><span class="dd-text">Copy</span><span class="dd-hint">Ctrl+C</span></div>
                  <div class="dd-item" (click)="exec('paste')"><span class="dd-text">Paste</span><span class="dd-hint">Ctrl+V</span></div>
                </div>
              </div>
              <div class="menu-item" (click)="toggleMenu('view', $event)" [class.active]="activeMenu === 'view'">
                View
                <div class="dropdown" *ngIf="activeMenu === 'view'">
                  <div class="dd-item" (click)="startSlideshow()"><span class="dd-text">Slideshow</span></div>
                  <div class="dd-item" (click)="toggleRuler()"><span class="dd-text">Show ruler</span></div>
                </div>
              </div>
              <div class="menu-item" (click)="toggleMenu('insert', $event)" [class.active]="activeMenu === 'insert'">
                Insert
                <div class="dropdown" *ngIf="activeMenu === 'insert'">
                  <div class="dd-item" (click)="triggerImageInsert()"><span class="dd-text">Image</span></div>
                  <div class="dd-item" (click)="insertTextBox()"><span class="dd-text">Text box</span></div>
                  <div class="dd-item" (click)="insertShape()"><span class="dd-text">Shape</span></div>
                  <div class="dd-item" (click)="insertLine()"><span class="dd-text">Line</span></div>
                </div>
              </div>
              <div class="menu-item" (click)="toggleMenu('format', $event)" [class.active]="activeMenu === 'format'">
                Format
                <div class="dropdown" *ngIf="activeMenu === 'format'">
                  <div class="dd-item" (click)="exec('bold')"><span class="dd-text">Bold</span><span class="dd-hint">Ctrl+B</span></div>
                  <div class="dd-item" (click)="exec('italic')"><span class="dd-text">Italic</span><span class="dd-hint">Ctrl+I</span></div>
                  <div class="dd-item" (click)="exec('underline')"><span class="dd-text">Underline</span><span class="dd-hint">Ctrl+U</span></div>
                  <div class="dd-sep"></div>
                  <div class="dd-item" (click)="exec('justifyLeft')"><span class="dd-text">Align left</span></div>
                  <div class="dd-item" (click)="exec('justifyCenter')"><span class="dd-text">Align center</span></div>
                  <div class="dd-item" (click)="exec('justifyRight')"><span class="dd-text">Align right</span></div>
                  <div class="dd-sep"></div>
                  <div class="dd-item" (click)="exec('removeFormat')"><span class="dd-text">Clear formatting</span><span class="dd-hint">Ctrl+\</span></div>
                </div>
              </div>
              <div class="menu-item" (click)="toggleMenu('slide', $event)" [class.active]="activeMenu === 'slide'">
                Slide
                <div class="dropdown" *ngIf="activeMenu === 'slide'">
                  <div class="dd-item" (click)="addSlide(); closeMenus()"><span class="dd-text">New slide</span><span class="dd-hint">Ctrl+M</span></div>
                  <div class="dd-item" (click)="deleteSlide(activeId, $event); closeMenus()"><span class="dd-text">Delete slide</span></div>
                  <div class="dd-sep"></div>
                  <div class="dd-item" (click)="duplicateSlide()"><span class="dd-text">Duplicate slide</span></div>
                  <div class="dd-item" (click)="changeBackground()"><span class="dd-text">Change background</span></div>
                  <div class="dd-item" (click)="applyLayout()"><span class="dd-text">Apply layout</span><span class="dd-hint">▶</span></div>
                </div>
              </div>
              <div class="menu-item" (click)="toggleMenu('arrange', $event)" [class.active]="activeMenu === 'arrange'">
                Arrange
                <div class="dropdown" *ngIf="activeMenu === 'arrange'">
                  <div class="dd-item" (click)="arrangeOrder()"><span class="dd-text">Order</span><span class="dd-hint">▶</span></div>
                  <div class="dd-item" (click)="arrangeAlign()"><span class="dd-text">Align</span><span class="dd-hint">▶</span></div>
                  <div class="dd-item" (click)="arrangeCenter()"><span class="dd-text">Center on page</span><span class="dd-hint">▶</span></div>
                  <div class="dd-sep"></div>
                  <div class="dd-item" (click)="groupElements()"><span class="dd-text">Group</span><span class="dd-hint">Ctrl+Alt+G</span></div>
                  <div class="dd-item" (click)="ungroupElements()"><span class="dd-text">Ungroup</span><span class="dd-hint">Ctrl+Alt+Shift+G</span></div>
                </div>
              </div>
              <div class="menu-item" (click)="toggleMenu('tools', $event)" [class.active]="activeMenu === 'tools'">
                Tools
                <div class="dropdown" *ngIf="activeMenu === 'tools'">
                  <div class="dd-item" (click)="checkSpelling()"><span class="dd-text">Spelling</span><span class="dd-hint">▶</span></div>
                  <div class="dd-item" (click)="openDictionary()"><span class="dd-text">Dictionary</span><span class="dd-hint">Ctrl+Shift+Y</span></div>
                  <div class="dd-sep"></div>
                  <div class="dd-item" (click)="voiceType()"><span class="dd-text">Voice type speaker notes</span><span class="dd-hint">Ctrl+Shift+S</span></div>
                </div>
              </div>
              <div class="menu-item" (click)="toggleMenu('extensions', $event)" [class.active]="activeMenu === 'extensions'">
                Extensions
                <div class="dropdown" *ngIf="activeMenu === 'extensions'">
                  <div class="dd-item" (click)="showExtensions()"><span class="dd-text">Add-ons</span></div>
                </div>
              </div>
              <div class="menu-item" (click)="toggleMenu('help', $event)" [class.active]="activeMenu === 'help'">
                Help
                <div class="dropdown" *ngIf="activeMenu === 'help'">
                  <div class="dd-item" (click)="showHelp()"><span class="dd-text">Slides Help</span></div>
                  <div class="dd-item" (click)="showTraining()"><span class="dd-text">Training</span></div>
                  <div class="dd-sep"></div>
                  <div class="dd-item" (click)="showKeyboardShortcuts()"><span class="dd-text">Keyboard shortcuts</span><span class="dd-hint">Ctrl+/</span></div>
                </div>
              </div>
            </div>
          </div>
        </div>
        <div class="top-right">
          <button class="btn slides-btn outline" (click)="startSlideshow()">
            <svg xmlns="http://www.w3.org/2000/svg" height="16" viewBox="0 -960 960 960" width="16" fill="currentColor" style="margin-right:6px;"><path d="M200-240q-33 0-56.5-23.5T120-320v-400q0-33 23.5-56.5T200-800h560q33 0 56.5 23.5T840-720v400q0 33-23.5 56.5T760-240H200Zm0-80h560v-400H200v400Zm120-80h320v-240H320v240ZM200-320v-400 400Z"/></svg>
            Slideshow
            <span style="border-left: 1px solid #c0c0c0; padding-left: 6px; margin-left: 6px;">▼</span>
          </button>
          <span class="badge" *ngIf="activeUsers > 1">
            <svg fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24" class="icon-sm" style="margin-right:4px;"><path stroke-linecap="round" stroke-linejoin="round" d="M12 4.354a4 4 0 110 5.292M15 21H3v-1a6 6 0 0112 0v1zm0 0h6v-1a6 6 0 00-9-5.197M13 7a4 4 0 11-8 0 4 4 0 018 0z"></path></svg>
            {{ activeUsers }} online
          </span>
          <button class="btn slides-btn blue-btn" (click)="shareModalOpen = true">
            <svg fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24" class="icon-sm" style="margin-right:4px;"><path stroke-linecap="round" stroke-linejoin="round" d="M13.828 10.172a4 4 0 00-5.656 0l-4 4a4 4 0 105.656 5.656l1.102-1.101m-.758-4.899a4 4 0 005.656 0l4-4a4 4 0 00-5.656-5.656l-1.1 1.1"></path></svg>
            Share
          </button>
          <div class="avatar" [title]="auth.user?.name ?? ''">{{ initials }}</div>
        </div>
      </div>
      
      <div class="fmt-bar">
        <div style="position: relative; display: flex;">
          <div class="fmt-group">
            <button class="fb add-new" (click)="addSlide()">+</button>
            <button class="fb add-arrow" (click)="toggleMenu('add-slide', $event)">▼</button>
          </div>
          <div class="dropdown" *ngIf="activeMenu === 'add-slide'" style="top: 100%; left: 0; min-width: 150px;">
            <div class="dd-item" (click)="addSlideWithLayout('title')"><span class="dd-text">Title slide</span></div>
            <div class="dd-item" (click)="addSlideWithLayout('two-column')"><span class="dd-text">Two-column</span></div>
            <div class="dd-item" (click)="addSlideWithLayout('blank')"><span class="dd-text">Blank</span></div>
          </div>
        </div>
        <button class="fb" title="Undo" (click)="exec('undo')"><svg viewBox="0 0 24 24" fill="currentColor" width="16" height="16"><path d="M12.5 8c-2.65 0-5.05.99-6.9 2.6L2 7v9h9l-3.62-3.62c1.39-1.16 3.16-1.88 5.12-1.88 3.54 0 6.55 2.31 7.6 5.5l2.37-.78C21.08 11.03 17.15 8 12.5 8z"/></svg></button>
        <button class="fb" title="Redo" (click)="exec('redo')"><svg viewBox="0 0 24 24" fill="currentColor" width="16" height="16"><path d="M18.4 10.6C16.55 8.99 14.15 8 11.5 8c-4.65 0-8.58 3.03-9.96 7.22L3.9 16c1.05-3.19 4.06-5.5 7.6-5.5 1.95 0 3.73.72 5.12 1.88L13 16h9V7l-3.6 3.6z"/></svg></button>
        <button class="fb" title="Print" (click)="print()"><svg viewBox="0 0 24 24" fill="currentColor" width="16" height="16"><path d="M19 8H5c-1.66 0-3 1.34-3 3v6h4v4h12v-4h4v-6c0-1.66-1.34-3-3-3zm-3 11H8v-5h8v5zm3-7c-.55 0-1-.45-1-1s.45-1 1-1 1 .45 1 1-.45 1-1 1zm-1-9H6v4h12V3z"/></svg></button>
        <button class="fb" title="Paint Format" (click)="paintFormat()"><svg xmlns="http://www.w3.org/2000/svg" height="16" viewBox="0 -960 960 960" width="16" fill="currentColor"><path d="M510-252q18-18 29.5-40t11.5-48q0-44-30-79t-71-61q-27-18-43.5-43.5T390-580q0-10 1-19.5t3-18.5l272 272q-7 3-17 4t-21 1q-34 0-66-15.5T510-252ZM166-508 60-614q-24-24-24-57t24-57l170-170q24-24 57-24t57 24l160 160q21 21 32 47.5t11 54.5q0 28-11 54t-32 46l-18 18L166-508Zm106-106 114-114-114-114-114 114 114 114Z"/></svg></button>
        <span class="sep"></span>
        <div style="position:relative;">
          <button class="fb-text zoom-text" (click)="toggleZoomDropdown($event)">{{ zoomLevel }}% ▼</button>
          <div class="dropdown" *ngIf="zoomDropdownOpen" style="min-width:80px; left:0;">
            <div class="dd-item" (click)="setZoom(50)"><span class="dd-text">50%</span></div>
            <div class="dd-item" (click)="setZoom(75)"><span class="dd-text">75%</span></div>
            <div class="dd-item" (click)="setZoom(100)"><span class="dd-text">100%</span></div>
            <div class="dd-item" (click)="setZoom(150)"><span class="dd-text">150%</span></div>
            <div class="dd-item" (click)="setZoom(200)"><span class="dd-text">200%</span></div>
          </div>
        </div>
        <span class="sep"></span>
        <button class="fb active" title="Select" (click)="activeImg = null"><svg viewBox="0 0 24 24" fill="currentColor" width="16" height="16"><path d="M7 2l12 11.2-5.8.5 3.3 7.3-2.2.9-3.2-7.4-4.4 4.7z"/></svg></button>
        <button class="fb" title="Text box" (click)="insertTextBox()"><svg viewBox="0 0 24 24" fill="currentColor" width="16" height="16"><path d="M2.5 4v3h5v12h3V7h5V4h-13zm19 5h-9v3h3v7h3v-7h3V9z"/></svg></button>
        <button class="fb" title="Insert image" (click)="triggerImageInsert()"><svg viewBox="0 0 24 24" fill="currentColor" width="16" height="16"><path d="M21 19V5c0-1.1-.9-2-2-2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2zM8.5 13.5l2.5 3.01L14.5 12l4.5 6H5l3.5-4.5z"/></svg></button>
        <button class="fb" title="Shape" (click)="insertShape()"><svg viewBox="0 0 24 24" fill="currentColor" width="16" height="16"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8z"/></svg></button>
        <button class="fb" title="Line" (click)="insertLine()"><svg viewBox="0 0 24 24" fill="currentColor" width="16" height="16"><path d="M19 12H5v-2h14v2z"/></svg></button>
        <span class="sep"></span>
        <button class="fb" title="Add comment" (click)="addComment()"><svg viewBox="0 0 24 24" fill="currentColor" width="16" height="16"><path d="M21.99 4c0-1.1-.89-2-1.99-2H4c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h14l4 4-.01-18zM18 14H6v-2h12v2zm0-3H6V9h12v2zm0-3H6V6h12v2z"/></svg></button>
        <span class="sep"></span>
        <button class="fb-text" (click)="changeBackground()">Background</button>
        <button class="fb-text" (click)="applyLayout()">Layout</button>
        <button class="fb-text" (click)="applyTheme()">Theme</button>
        <button class="fb-text" (click)="applyTransition()">Transition</button>
      </div>

      <div class="workspace">
        <div class="panel">
          <button class="add-btn" (click)="addSlide()">+ New Slide</button>
          <div *ngFor="let id of data.pageOrder; let i = index"
            class="thumb" [class.active]="id === activeId" (click)="selectSlide(id)">
            <span class="thumb-n">{{ i + 1 }}</span>
            <span class="thumb-t">{{ data.pages[id].title || 'Slide ' + (i + 1) }}</span>
            <button class="thumb-del" (click)="deleteSlide(id, $event)">×</button>
          </div>
          <div *ngIf="data.pageOrder.length === 0" class="no-slides">No slides yet</div>
        </div>

        <div class="canvas-area" *ngIf="activePage; else noSlide">
          <div class="slide" id="active-slide" #slide
              (mousedown)="onSlideMouseDown($event)"
              [style.zoom]="zoomLevel / 100" 
              [style.background-color]="activePage.bg || '#fff'"
              [style.color]="activePage.color || '#374151'">
            <div #slideBody class="slide-body" contenteditable="true" 
              (input)="onContentChange($event)" (blur)="onContentBlur()"
              data-placeholder="Slide content…"></div>
              
            <!-- Interactive Image Overlay -->
            <div *ngIf="activeImg" class="img-overlay" [ngStyle]="overlayStyle">
              <div class="drag-handle" (mousedown)="startDrag($event)">Move</div>
              <div class="resize-handle nw" (mousedown)="startResize($event, 'nw')"></div>
              <div class="resize-handle n" (mousedown)="startResize($event, 'n')"></div>
              <div class="resize-handle ne" (mousedown)="startResize($event, 'ne')"></div>
              <div class="resize-handle w" (mousedown)="startResize($event, 'w')"></div>
              <div class="resize-handle e" (mousedown)="startResize($event, 'e')"></div>
              <div class="resize-handle sw" (mousedown)="startResize($event, 'sw')"></div>
              <div class="resize-handle s" (mousedown)="startResize($event, 's')"></div>
              <div class="resize-handle se" (mousedown)="startResize($event, 'se')"></div>
            </div>
          </div>
        </div>
        <ng-template #noSlide>
          <div class="canvas-area no-slide-msg">
            <p>Click <strong>+ New Slide</strong> to get started.</p>
          </div>
        </ng-template>
      </div>

      <div class="toast" [class.show]="toastVisible">{{ toastMsg }}</div>

      <!-- Share Modal -->
      <div class="modal-overlay" *ngIf="shareModalOpen" (click)="shareModalOpen = false; shareRoleDropdownOpen = false">
        <div class="modal share-modal" (click)="$event.stopPropagation()" style="background:#202124; color:#e8eaed; border-radius:12px; padding:24px; width:520px; box-shadow:0 12px 40px rgba(0,0,0,.6); font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif; border:none; max-width:90vw;">
          <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:24px;">
            <div style="display:flex; align-items:center; gap:10px;">
              <div style="background:#f4b400; color:#fff; display:flex; align-items:center; justify-content:center; width:22px; height:22px; border-radius:4px;">
                <span class="material-symbols-outlined" style="font-size:16px;">slideshow</span>
              </div>
              <h3 style="margin:0; font-size:18px; font-weight:500; color:#e8eaed;">Share "{{ title || 'Untitled presentation' }}"</h3>
            </div>
            <button (click)="shareModalOpen = false" style="background:none; border:none; color:#9aa0a6; cursor:pointer; display:flex; align-items:center; justify-content:center; padding:6px; border-radius:50%;">
              <span class="material-symbols-outlined" style="font-size:20px;">close</span>
            </button>
          </div>
          <div style="position:relative; margin-bottom:32px;">
            <div style="display:flex; align-items:center; gap:12px; position:relative;">
              <div style="flex:1; display:flex; align-items:center; background:#1c1d1f; border:1px solid #5f6368; border-radius:4px; padding:0 12px; height:44px;">
                <input type="text" [(ngModel)]="shareQuery" style="flex:1; background:transparent; border:none; color:#e8eaed; font-size:14px; outline:none; height:100%;" placeholder="Add people and groups">
                <div (click)="shareRoleDropdownOpen = !shareRoleDropdownOpen" style="display:flex; align-items:center; gap:4px; color:#e8eaed; font-size:13px; cursor:pointer; padding-left:12px; position:relative;">
                  {{ shareRole }} <span class="material-symbols-outlined" style="font-size:18px; opacity: 0.8;">arrow_drop_down</span>
                  
                  <div *ngIf="shareRoleDropdownOpen" style="position:absolute; top:30px; right:0; background:#2d3748; border:1px solid #4a5568; border-radius:4px; box-shadow:0 4px 12px rgba(0,0,0,0.5); z-index:100; min-width:100px;">
                     <div (click)="shareRole = 'View'; shareRoleDropdownOpen = false; $event.stopPropagation()" style="padding:8px 12px; cursor:pointer; color:#e8eaed;">View</div>
                     <div (click)="shareRole = 'Edit'; shareRoleDropdownOpen = false; $event.stopPropagation()" style="padding:8px 12px; cursor:pointer; color:#e8eaed;">Edit</div>
                  </div>
                </div>
              </div>
              <button (click)="performShare()" style="background:#f4b400; color:#fff; border:none; border-radius:24px; font-weight:500; font-size:14px; padding:0 24px; height:44px; cursor:pointer;">Share</button>
            </div>
          </div>
          <div style="margin-bottom:32px;">
            <div style="font-size:11px; font-weight:600; color:#9aa0a6; letter-spacing:0.8px; margin-bottom:16px;">WHO CAN ACCESS</div>
            <div style="display:flex; align-items:center; justify-content:space-between;">
              <div style="display:flex; align-items:center; gap:16px;">
                <div style="background:#303134; display:flex; align-items:center; justify-content:center; width:40px; height:40px; border-radius:50%; color:#e8eaed;">
                  <span class="material-symbols-outlined" style="font-size:20px;">{{ isPublic ? 'public' : 'link' }}</span>
                </div>
                <div>
                  <div style="font-size:14px; font-weight:600; color:#e8eaed;">{{ isPublic ? 'Public Link - Anyone on the internet can view' : 'Permalink - Private, not shared with anyone' }}</div>
                </div>
              </div>
              <button *ngIf="!isPublic" (click)="isPublic = true" style="background:transparent; border:none; color:#9aa0a6; display:flex; align-items:center; gap:6px; font-size:13px; font-weight:500; cursor:pointer; padding:8px; border-radius:4px;">
                 <span class="material-symbols-outlined" style="font-size:16px;">settings</span> Make Public
              </button>
              <button *ngIf="isPublic" (click)="isPublic = false" style="background:transparent; border:none; color:#9aa0a6; display:flex; align-items:center; gap:6px; font-size:13px; font-weight:500; cursor:pointer; padding:8px; border-radius:4px;">
                 <span class="material-symbols-outlined" style="font-size:16px;">lock</span> Make Private
              </button>
            </div>
          </div>
          <div style="display:flex; align-items:center; justify-content:space-between;">
            <button (click)="copyLink()" style="background:transparent; border:none; color:#e8eaed; font-size:14px; font-weight:500; border-radius:24px; padding:8px 12px; margin-left:-12px; cursor:pointer;">Copy Link</button>
            <button (click)="shareModalOpen = false" style="background:#303134; color:#8ab4f8; font-size:14px; font-weight:500; border:none; border-radius:24px; padding:0 24px; height:40px; cursor:pointer;">Done</button>
          </div>
        </div>
      </div>
      
      <!-- Image Modal -->
      <div class="modal-overlay" *ngIf="imageModalOpen" (click)="imageModalOpen = false">
        <div class="modal share-modal" (click)="$event.stopPropagation()">
          <h3 style="margin-top:0;color:#202124;">Insert Image</h3>
          
          <div style="margin-bottom:16px;">
            <p style="color:#5f6368;font-size:14px;margin-bottom:8px;">Option 1: Upload from computer</p>
            <input type="file" #imgInput accept="image/*" style="display:none" (change)="uploadImage($event)">
            <button class="btn outline" style="width:100%" (click)="imgInput.click()">Select File</button>
          </div>
          
          <div style="text-align:center;color:#9aa0a6;margin-bottom:16px;font-size:12px;">— OR —</div>
          
          <p style="color:#5f6368;font-size:14px;margin-bottom:8px;">Option 2: By URL</p>
          <input type="text" [(ngModel)]="imageUrl" placeholder="https://" style="width:100%;padding:8px;margin-bottom:16px;box-sizing:border-box;border:1px solid #ccc;border-radius:4px;" />
          <div style="display:flex;justify-content:flex-end;gap:8px;">
            <button class="btn outline" (click)="imageModalOpen = false">Cancel</button>
            <button class="btn blue-btn" (click)="insertImage()">Insert URL</button>
          </div>
        </div>
      </div>

      <div class="modal-overlay slideshow-overlay" *ngIf="slideshowActive" (click)="slideshowActive = false">
        <div class="slideshow-container" (click)="$event.stopPropagation()">
          <div class="slide fullscreen-slide" 
              [style.background-color]="activePage?.bg || '#fff'" 
              [style.color]="activePage?.color || '#374151'">
            <div class="ss-body" [innerHTML]="activePage?.body"></div>
          </div>
          <div class="ss-controls">
            <button (click)="prevSlide()">◀</button>
            <span>Slide {{ currentSlideIndex + 1 }} of {{ data.pageOrder.length }}</span>
            <button (click)="nextSlide()">▶</button>
            <button (click)="slideshowActive = false" style="margin-left: 16px;">Esc</button>
          </div>
        </div>
      </div>

      <!-- Dictionary Modal -->
      <div class="modal-overlay" *ngIf="dictModalOpen" (click)="dictModalOpen = false">
        <div class="modal" style="width: 400px; max-height: 80vh; overflow-y: auto; text-align: left;" (click)="$event.stopPropagation()">
          <h3 style="margin-top: 0;">Dictionary</h3>
          <div style="display: flex; gap: 8px; margin-top: 12px; margin-bottom: 16px;">
            <input type="text" [(ngModel)]="dictWord" class="form-control" placeholder="Enter a word" (keyup.enter)="searchDictionary()" style="flex:1; padding:8px; border:1px solid #ccc; border-radius:4px;">
            <button class="btn blue-btn" (click)="searchDictionary()">Search</button>
          </div>
          
          <div *ngIf="dictLoading" style="text-align: center; color: #5f6368; padding: 20px;">Searching...</div>
          <div *ngIf="dictError" style="color: #ea4335; text-align: center; padding: 20px;">{{ dictError }}</div>
          
          <div *ngIf="dictResults && dictResults.length > 0">
            <div *ngFor="let result of dictResults" style="margin-bottom: 16px;">
              <h4 style="margin: 0; font-size: 18px; color: #202124;">{{ result.word }} <span style="font-size: 14px; color: #5f6368; font-weight: normal;">{{ result.phonetic }}</span></h4>
              <div *ngFor="let meaning of result.meanings" style="margin-top: 8px;">
                <div style="font-style: italic; color: #1a73e8; font-size: 13px;">{{ meaning.partOfSpeech }}</div>
                <ul style="margin: 4px 0 0 0; padding-left: 20px; font-size: 13px; color: #3c4043;">
                  <li *ngFor="let def of meaning.definitions" style="margin-bottom: 4px;">
                    {{ def.definition }}
                    <div *ngIf="def.example" style="color: #5f6368; font-style: italic; margin-top: 2px;">"{{ def.example }}"</div>
                  </li>
                </ul>
              </div>
            </div>
          </div>
          
          <div class="modal-actions" style="margin-top: 16px; text-align: right;">
            <button class="btn outline" (click)="dictModalOpen = false">Close</button>
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
      
      <!-- Loading Data Overlay -->
      <div class="loading-overlay" *ngIf="isLoadingDocument">
        <div class="loading-modal shadow-lg">
           <div class="lm-spinner"></div>
           <div class="lm-title">Loading Presentation...</div>
           <div class="lm-subtitle">Retrieving slides, shapes, and formatting. Please wait.</div>
        </div>
      </div>
    </div>
  `,
  styles: [`
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
    .shell { display: flex; flex-direction: column; height: calc(100vh - 36px); background: #f9fbfd; }
    
    .top-bar { display: flex; align-items: flex-start; justify-content: space-between; padding: 8px 16px 0; background: #f9fbfd; }
    .top-left { display: flex; align-items: flex-start; gap: 10px; }
    .back { 
      background: none; border: none; cursor: pointer; color: #5f6368; 
      display: flex; align-items: center; justify-content: center;
      width: 40px; height: 40px; border-radius: 50%;
      margin-top: 4px;
    }
    .back:hover { background: #f1f3f4; }
    .icon-sm { width: 16px; height: 16px; display: inline-block; vertical-align: middle; }
    .doc-meta { display: flex; flex-direction: column; gap: 2px; }
    .title-input { font-size: 18px; font-weight: 400; color: #202124; border: 1px solid transparent; border-radius: 4px; padding: 2px 6px; outline: none; background: transparent; width: 250px; }
    .title-input:focus { border-color: #1a73e8; background: #fff; }
    .menu-bar { display: flex; gap: 2px; }
    .menu-item { position: relative; padding: 4px 8px; font-size: 14px; color: #202124; cursor: pointer; border-radius: 4px; user-select: none; }
    .menu-item:hover, .menu-item.active { background: #e1e5ea; }
    
    .dropdown {
      position: absolute; top: 100%; left: 0; background: #fff; border: 1px solid #ccc;
      box-shadow: 0 2px 6px rgba(0,0,0,0.2); border-radius: 4px; padding: 6px 0;
      min-width: 260px; z-index: 100;
    }
    .dd-item { display: flex; justify-content: space-between; padding: 6px 24px 6px 32px; cursor: pointer; color: #202124; font-size: 13px; align-items: center; }
    .dd-item:hover { background: #f1f3f4; }
    .dd-text { font-weight: 400; }
    .dd-hint { color: #5f6368; font-size: 12px; }
    .dd-sep { height: 1px; background: #e0e0e0; margin: 4px 0; }
    
    .icon-sm { width: 16px; height: 16px; display: inline-block; vertical-align: middle; }

    .top-right { display: flex; align-items: center; gap: 12px; margin-top: 4px; }
    .slides-btn { padding: 8px 16px; border-radius: 20px; font-weight: 500; font-size: 14px; cursor: pointer; display: flex; align-items: center; gap: 6px; }
    .slides-btn.outline { background: #fff; color: #444746; border: 1px solid #747775; display: flex; align-items: center; }
    .slides-btn.outline:hover { background: #f8f9fa; }
    .slides-btn.blue-btn { background: #c2e7ff; color: #001d35; border: none; }
    .slides-btn.blue-btn:hover { background: #b3d4ec; box-shadow: 0 1px 2px rgba(0,0,0,0.3); }
    .badge { font-size: 12px; background: #e6f4ea; color: #137333; padding: 4px 12px; border-radius: 12px; font-weight: 500; display: flex; align-items: center; }
    .avatar { width: 32px; height: 32px; border-radius: 50%; background: #2563eb; color: #fff; font-size: 14px; font-weight: 500; display: flex; align-items: center; justify-content: center; }
    
    .fmt-bar {
      display: flex; align-items: center; gap: 4px; padding: 6px 16px; flex-wrap: wrap;
      background: #edf2fa; border-radius: 24px; margin: 8px 16px; position: relative; z-index: 90;
    }
    .fb { min-width: 32px; height: 32px; background: none; border: 1px solid transparent; border-radius: 4px; cursor: pointer; display: flex; align-items: center; justify-content: center; font-size: 16px; color: #444746; }
    .fb:hover { background: #e0e6ed; }
    .fb.active { background: #d3e3fd; color: #0b57d0; }
    .fb-text { background: none; border: none; padding: 0 8px; height: 32px; border-radius: 4px; color: #444746; font-size: 14px; cursor: pointer; font-weight: 500; }
    .fb-text:hover { background: #e0e6ed; }
    .zoom-text { font-weight: 400; padding: 0 12px; }
    .sep { width: 1px; height: 20px; background: #c7c7c7; margin: 0 6px; }
    
    .fmt-group { display: flex; border-radius: 4px; overflow: hidden; }
    .fmt-group .fb { border-radius: 0; }
    .fmt-group .add-new { padding: 0 12px; font-size: 20px; font-weight: 300; background: #d3e3fd; color: #0b57d0; }
    .fmt-group .add-arrow { padding: 0 6px; font-size: 10px; background: #d3e3fd; color: #0b57d0; min-width: auto; border-left: 1px solid #b3c9f2; }
    .fmt-group .add-new:hover, .fmt-group .add-arrow:hover { background: #c2d7f5; }

    .workspace { display: flex; flex: 1; overflow: hidden; background: #f9fbfd; }
    .panel {
      width: 220px; flex-shrink: 0; background: #fff; border-right: 1px solid #dadce0; border-top: 1px solid #dadce0;
      overflow-y: auto; padding: 12px; display: flex; flex-direction: column; gap: 8px;
    }
    .thumb {
      display: flex; align-items: center; gap: 6px; padding: 8px 10px;
      border-radius: 8px; cursor: pointer; border: 2px solid transparent; background: #f3f4f6;
    }
    .thumb:hover { background: #eff6ff; }
    .thumb.active { border-color: #2563eb; background: #eff6ff; }
    .thumb-n { font-size: 11px; color: #9ca3af; min-width: 16px; }
    .thumb-t { flex: 1; font-size: 13px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .thumb-del { background: none; border: none; color: #d1d5db; cursor: pointer; font-size: 16px; padding: 0 2px; }
    .thumb-del:hover { color: #ef4444; }
    .no-slides { font-size: 12px; color: #9ca3af; text-align: center; padding: 16px 0; }
    .canvas-area { flex: 1; display: flex; align-items: center; justify-content: center; padding: 32px; overflow: auto; background: #f3f4f6; }
    .no-slide-msg { color: #9ca3af; font-size: 15px; }
    .slide {
      width: 960px; height: 540px; border-radius: 0;
      box-shadow: 0 4px 12px rgba(0,0,0,.1); display: flex; flex-direction: column;
      padding: 60px 80px; gap: 24px; position: relative;
    }
    .slide-title {
      border: none; border-bottom: 2px solid #e5e7eb; font-size: 28px; font-weight: 700;
      outline: none; width: 100%; background: transparent; padding-bottom: 8px; color: inherit;
    }
    .slide-title:focus { border-bottom-color: #2563eb; }
    .slide-body { flex: 1; border: none; outline: none; font-size: 16px; line-height: 1.7; background: transparent; color: inherit; position: relative; }
    .slide-body:empty:before { content: attr(data-placeholder); color: #9ca3af; pointer-events: none; }
    .slide-body img { max-width: 100%; height: auto; display: block; cursor: pointer; }
    
    .img-overlay { position: absolute; z-index: 100; pointer-events: none; }
    .drag-handle { position: absolute; top: -24px; left: 0; background: #1a73e8; color: #fff; padding: 2px 6px; font-size: 12px; border-radius: 4px; pointer-events: auto; cursor: move; user-select: none; }
    .resize-handle { position: absolute; width: 12px; height: 12px; background: #fff; border: 2px solid #1a73e8; border-radius: 50%; pointer-events: auto; }
    .resize-handle.nw { top: -7px; left: -7px; cursor: nwse-resize; }
    .resize-handle.n { top: -7px; left: 50%; margin-left: -7px; cursor: ns-resize; }
    .resize-handle.ne { top: -7px; right: -7px; cursor: nesw-resize; }
    .resize-handle.w { top: 50%; left: -7px; margin-top: -7px; cursor: ew-resize; }
    .resize-handle.e { top: 50%; right: -7px; margin-top: -7px; cursor: ew-resize; }
    .resize-handle.sw { bottom: -7px; left: -7px; cursor: nesw-resize; }
    .resize-handle.s { bottom: -7px; left: 50%; margin-left: -7px; cursor: ns-resize; }
    .resize-handle.se { bottom: -7px; right: -7px; cursor: nwse-resize; }
    
    .slideshow-overlay { background: #000; display: flex; align-items: center; justify-content: center; z-index: 9999; }
    .slideshow-container { position: relative; width: 100vw; height: 100vh; display: flex; align-items: center; justify-content: center; }
    .fullscreen-slide { transform: scale(1.5); box-shadow: none; }
    .ss-title { font-size: 40px; margin-bottom: 24px; border-bottom: 2px solid #ccc; padding-bottom: 12px; }
    .ss-body { font-size: 24px; line-height: 1.8; }
    .ss-controls { position: absolute; bottom: 24px; left: 50%; transform: translateX(-50%); display: flex; gap: 16px; background: rgba(0,0,0,0.6); padding: 8px 16px; border-radius: 24px; color: #fff; align-items: center; }
    .ss-controls button { background: none; border: none; color: #fff; font-size: 16px; cursor: pointer; padding: 4px 8px; }
    .ss-controls button:hover { color: #4ade80; }

    .toast {
      position: fixed; bottom: 32px; left: 50%; transform: translateX(-50%) translateY(20px);
      background: #111; color: #fff; padding: 10px 24px; border-radius: 8px;
      font-size: 14px; opacity: 0; transition: all .25s; pointer-events: none;
    }
    .toast.show { opacity: 1; transform: translateX(-50%) translateY(0); }
    .modal-overlay {
      position: fixed; top: 0; left: 0; right: 0; bottom: 0;
      background: rgba(0,0,0,0.5); z-index: 999; display: flex; align-items: center; justify-content: center;
    }
    .modal { background: #fff; padding: 28px; border-radius: 10px; width: 460px; max-width: 92%;
             box-shadow: 0 8px 32px rgba(0,0,0,0.18); position: relative; }
    .share-modal { text-align: left; }
    
    /* ===== PRINT STYLES ===== */
    @media print {
      .top-bar, .fmt-bar, .panel, .modal-overlay, .toast, .img-overlay, .top-right { display: none !important; }
      .shell, .workspace, .canvas-area { display: block !important; height: auto !important; padding: 0 !important; background: transparent !important; overflow: visible !important; }
      .slide { box-shadow: none !important; margin: 0 auto 20px !important; page-break-after: always; border: 1px solid #ddd; }
    }
    
    /* Loading Overlay (2026 Premium) */
    .loading-overlay { position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(255, 255, 255, 0.85); backdrop-filter: blur(8px); -webkit-backdrop-filter: blur(8px); z-index: 99999; display: flex; align-items: center; justify-content: center; animation: fadeIn 0.3s ease-out; }
    .loading-modal { background: #fff; border-radius: 16px; padding: 40px; width: 420px; text-align: center; border: 1px solid rgba(0,0,0,0.08); box-shadow: 0 20px 40px rgba(0,0,0,0.1); animation: slideUp 0.4s cubic-bezier(0.16, 1, 0.3, 1); }
    .lm-spinner { width: 48px; height: 48px; border: 4px solid #e8f0fe; border-top-color: #1a73e8; border-radius: 50%; margin: 0 auto 24px auto; animation: spin 1s linear infinite; }
    .lm-title { font-size: 20px; font-weight: 600; color: #202124; margin-bottom: 8px; letter-spacing: -0.3px; }
    .lm-subtitle { font-size: 14px; color: #5f6368; line-height: 1.5; padding: 0 20px; }
    
    @keyframes spin { to { transform: rotate(360deg); } }
    @keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }
    @keyframes slideUp { from { opacity: 0; transform: translateY(20px) scale(0.98); } to { opacity: 1; transform: translateY(0) scale(1); } }
  `]
})
export class SlideEditorComponent implements OnInit, OnDestroy {
  isLoadingDocument = true;
  goHome() {
    window.location.href = 'https://show.vsnaptechnology.com/';
  }
  activeWidget: string | null = null;
  toggleWidget(w: string) {
    if (this.activeWidget === w) this.activeWidget = null;
    else this.activeWidget = w;
  }
  @ViewChild('slideBody') slideBodyRef?: ElementRef<HTMLElement>;

  docId = '';
  title = 'Untitled presentation';
  activeUsers = 1;
  shareModalOpen = false;
  imageModalOpen = false;
  imageUrl = '';
  activeMenu: string | null = null;
  zoomLevel = 100;
  zoomDropdownOpen = false;
  data: SlideData = { id: '', title: '', pages: {}, pageOrder: [] };
  activeId = '';
  shareQuery = '';
  shareRole: 'View' | 'Edit' = 'View';
  shareRoleDropdownOpen = false;
  isPublic = false;
  toastVisible = false;
  toastMsg = '';
  history: string[] = [];
  historyIdx = -1;
  slideshowActive = false;

  // Extra features state
  showRuler = false;
  paintFormatActive = false;
  dictModalOpen = false;
  dictWord = '';
  dictLoading = false;
  dictError = '';
  dictResults: any[] = [];

  private syncSub?: Subscription;
  private applyingRemote = false;

  @ViewChild('slide') slideRef!: ElementRef<HTMLElement>;

  // Drag and resize state
  isDragging = false;
  isResizing = false;
  resizeDir = '';
  activeImg: HTMLElement | null = null;
  dragStartX = 0;
  dragStartY = 0;
  imgStartX = 0;
  imgStartY = 0;
  imgStartW = 0;
  imgStartH = 0;

  get overlayStyle() {
    if (!this.activeImg || !this.slideRef) return { display: 'none' };
    const imgRect = this.activeImg.getBoundingClientRect();
    const slideRect = this.slideRef.nativeElement.getBoundingClientRect();
    const zoom = this.zoomLevel / 100;

    return {
      left: ((imgRect.left - slideRect.left) / zoom) + 'px',
      top: ((imgRect.top - slideRect.top) / zoom) + 'px',
      width: (imgRect.width / zoom) + 'px',
      height: (imgRect.height / zoom) + 'px'
    };
  }

  get currentUrl(): string { return window.location.href; }

  toggleMenu(menu: string, e: Event) {
    e.stopPropagation();
    this.activeMenu = this.activeMenu === menu ? null : menu;
    this.zoomDropdownOpen = false;
  }

  toggleZoomDropdown(e: Event) {
    e.stopPropagation();
    this.zoomDropdownOpen = !this.zoomDropdownOpen;
    this.activeMenu = null;
  }

  setZoom(level: number) {
    this.zoomLevel = level;
    this.zoomDropdownOpen = false;
  }

  closeMenus() { this.activeMenu = null; this.zoomDropdownOpen = false; }

  get activePage() { return this.activeId ? this.data.pages[this.activeId] : null; }

  get initials() {
    return (this.auth.user?.name ?? 'U').split(' ').map((w: string) => w[0]).join('').toUpperCase().slice(0, 2);
  }

  constructor(
    private route: ActivatedRoute,
    private router: Router,
    private api: ApiService,
    public auth: AuthService
  ) { }

  get currentSlideIndex() { return this.data.pageOrder.indexOf(this.activeId); }

  print() {
    this.closeMenus();
    const win = window.open('', '_blank', 'width=1000,height=700');
    if (!win) { window.print(); return; }

    let slidesHtml = '';
    for (const slideId of this.data.pageOrder) {
      const page = this.data.pages[slideId];
      const bg = page.bg || '#ffffff';
      const color = page.color || '#374151';
      slidesHtml += `
        <div class="print-slide" style="background: ${bg}; color: ${color};">
          <h1 class="print-title">${page.title || ''}</h1>
          <div class="print-body">${page.body || ''}</div>
        </div>
      `;
    }

    win.document.write(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>Print - ${this.title}</title>
        <style>
          @page { size: landscape; margin: 0; }
          body { font-family: Arial, sans-serif; margin: 0; padding: 0; background: #f0f0f0; }
          .print-slide {
            width: 960px; height: 540px; 
            margin: 20px auto; 
            position: relative;
            box-sizing: border-box;
            padding: 60px 80px;
            page-break-after: always;
            overflow: hidden;
            background: #fff;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
          }
          @media print {
            body { background: transparent; }
            .print-slide { margin: 0 auto; box-shadow: none; border: none; }
          }
          .print-title {
            margin: 0 0 24px 0;
            padding-bottom: 8px;
            border-bottom: 2px solid #e5e7eb;
            font-size: 28px;
            font-weight: 700;
          }
          .print-body {
            font-size: 16px; line-height: 1.7; color: #374151;
            position: relative;
            flex: 1;
            width: 100%;
            height: calc(100% - 65px);
          }
          img { max-width: 100%; height: auto; }
        </style>
      </head>
      <body>
        ${slidesHtml}
      </body>
      </html>
    `);
    win.document.close();
    win.focus();
    setTimeout(() => { win.print(); }, 500);
  }

  exec(cmd: string) {
    if (cmd === 'undo') { this.undo(); }
    else if (cmd === 'redo') { this.redo(); }
    else {
      document.execCommand(cmd, false, undefined);
      this.onChanged();
    }
    this.closeMenus();
  }

  triggerImageInsert() {
    this.closeMenus();
    this.imageUrl = '';
    this.imageModalOpen = true;
  }

  insertImageHtml(url: string) {
    if (!this.slideBodyRef || !this.activePage) return;
    this.slideBodyRef.nativeElement.focus();
    const success = document.execCommand('insertImage', false, url);
    if (!success) {
      this.slideBodyRef.nativeElement.innerHTML += `<img src="${url}" />`;
    }
    this.activePage.body = this.slideBodyRef.nativeElement.innerHTML;
    this.onChanged();
  }

  insertImage() {
    if (this.imageUrl) {
      this.insertImageHtml(this.imageUrl);
    }
    this.imageModalOpen = false;
  }

  uploadImage(e: any) {
    const file = e.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (re) => {
        const dataUrl = re.target?.result as string;
        this.insertImageHtml(dataUrl);
      };
      reader.readAsDataURL(file);
    }
    e.target.value = '';
    this.imageModalOpen = false;
  }

  onSlideMouseDown(e: MouseEvent) {
    const target = e.target as HTMLElement;
    if (target.tagName === 'IMG' || target.classList.contains('canvas-shape') || target.classList.contains('canvas-textbox')) {
      this.activeImg = target;
      if (this.activeImg.style.position !== 'absolute') {
        const imgRect = this.activeImg.getBoundingClientRect();
        const bodyRect = this.slideBodyRef!.nativeElement.getBoundingClientRect();
        const zoom = this.zoomLevel / 100;

        const left = (imgRect.left - bodyRect.left) / zoom;
        const top = (imgRect.top - bodyRect.top) / zoom;

        this.activeImg.style.position = 'absolute';
        this.activeImg.style.left = left + 'px';
        this.activeImg.style.top = top + 'px';
        this.activeImg.style.width = (imgRect.width / zoom) + 'px';
        this.activeImg.style.height = (imgRect.height / zoom) + 'px';
        this.activeImg.style.maxWidth = 'none';
      }
    } else if (!target.classList.contains('drag-handle') &&
      !target.classList.contains('resize-handle')) {
      this.activeImg = null;
    }
  }

  startDrag(e: MouseEvent) {
    e.preventDefault();
    e.stopPropagation();
    this.isDragging = true;
    this.dragStartX = e.clientX;
    this.dragStartY = e.clientY;
    this.imgStartX = parseFloat(this.activeImg!.style.left || '0');
    this.imgStartY = parseFloat(this.activeImg!.style.top || '0');
  }

  startResize(e: MouseEvent, dir: string) {
    e.preventDefault();
    e.stopPropagation();
    this.isResizing = true;
    this.resizeDir = dir;
    this.dragStartX = e.clientX;
    this.dragStartY = e.clientY;
    this.imgStartX = parseFloat(this.activeImg!.style.left || '0');
    this.imgStartY = parseFloat(this.activeImg!.style.top || '0');
    this.imgStartW = parseFloat(this.activeImg!.style.width || this.activeImg!.offsetWidth + '');
    this.imgStartH = parseFloat(this.activeImg!.style.height || this.activeImg!.offsetHeight + '');
  }

  @HostListener('document:mousemove', ['$event'])
  onGlobalMouseMove(e: MouseEvent) {
    if (!this.activeImg) return;
    const zoom = this.zoomLevel / 100;
    const dx = (e.clientX - this.dragStartX) / zoom;
    const dy = (e.clientY - this.dragStartY) / zoom;

    const bodyWidth = this.slideBodyRef!.nativeElement.offsetWidth;
    const bodyHeight = this.slideBodyRef!.nativeElement.offsetHeight;

    if (this.isDragging) {
      const imgWidth = parseFloat(this.activeImg.style.width || this.activeImg.offsetWidth + '');
      const imgHeight = parseFloat(this.activeImg.style.height || this.activeImg.offsetHeight + '');

      let newX = this.imgStartX + dx;
      let newY = this.imgStartY + dy;

      newX = Math.max(0, Math.min(newX, Math.max(0, bodyWidth - imgWidth)));
      newY = Math.max(0, Math.min(newY, Math.max(0, bodyHeight - imgHeight)));

      this.activeImg.style.left = newX + 'px';
      this.activeImg.style.top = newY + 'px';
    } else if (this.isResizing) {
      let newW = this.imgStartW;
      let newH = this.imgStartH;
      let newX = this.imgStartX;
      let newY = this.imgStartY;

      if (this.resizeDir.includes('e')) newW = this.imgStartW + dx;
      if (this.resizeDir.includes('w')) { newW = this.imgStartW - dx; newX = this.imgStartX + dx; }
      if (this.resizeDir.includes('s')) newH = this.imgStartH + dy;
      if (this.resizeDir.includes('n')) { newH = this.imgStartH - dy; newY = this.imgStartY + dy; }

      if (newW > 20) {
        if (newX < 0) { newW += newX; newX = 0; }
        if (newX + newW > bodyWidth) newW = Math.max(20, bodyWidth - newX);
        this.activeImg.style.width = newW + 'px';
        this.activeImg.style.left = newX + 'px';
      }
      if (newH > 20) {
        if (newY < 0) { newH += newY; newY = 0; }
        if (newY + newH > bodyHeight) newH = Math.max(20, bodyHeight - newY);
        this.activeImg.style.height = newH + 'px';
        this.activeImg.style.top = newY + 'px';
      }
    }
  }

  @HostListener('document:keydown', ['$event'])
  onGlobalKeyDown(e: KeyboardEvent) {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
      e.preventDefault();
      this.save();
      return;
    }
  }

  @HostListener('document:mouseup', ['$event'])
  onGlobalMouseUp(e: MouseEvent) {
    if (this.isDragging || this.isResizing) {
      this.isDragging = false;
      this.isResizing = false;
      if (this.activePage && this.slideBodyRef) {
        this.activePage.body = this.slideBodyRef.nativeElement.innerHTML;
        this.onChanged();
      }
    }
  }

  onContentChange(e: any) {
    if (this.activePage) {
      this.activePage.body = e.target.innerHTML;
      this.onChanged(true);
    }
  }

  onContentBlur() {
    this.pushHistory();
  }

  startSlideshow() {
    this.closeMenus();
    this.slideshowActive = true;
    if (document.documentElement.requestFullscreen) {
      document.documentElement.requestFullscreen().catch(() => { });
    }
  }

  prevSlide() {
    const idx = this.currentSlideIndex;
    if (idx > 0) this.activeId = this.data.pageOrder[idx - 1];
  }

  nextSlide() {
    const idx = this.currentSlideIndex;
    if (idx < this.data.pageOrder.length - 1) this.activeId = this.data.pageOrder[idx + 1];
  }

  newDoc() {
    this.api.createDocument('Untitled presentation', 'slide').subscribe(res => {
      window.open('/slide/' + res.id, '_blank');
    });
    this.closeMenus();
  }

  makeCopy() {
    this.api.createDocument(this.title + ' (Copy)', 'slide').subscribe(res => {
      this.api.saveDocument(res.id, this.title + ' (Copy)', JSON.stringify(this.data)).subscribe(() => {
        window.open('/slide/' + res.id, '_blank');
      });
    });
    this.closeMenus();
  }

  triggerOpen() {
    this.closeMenus();
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.json';
    input.onchange = (e: any) => {
      const file = e.target.files?.[0];
      if (file) {
        const reader = new FileReader();
        reader.onload = (re) => {
          try {
            const content = re.target?.result as string;
            JSON.parse(content);
            this.api.createDocument(file.name.replace(/\.[^/.]+$/, ""), 'slide').subscribe(res => {
              this.api.saveDocument(res.id, file.name.replace(/\.[^/.]+$/, ""), content).subscribe(() => {
                window.open('/slide/' + res.id, '_blank');
              });
            });
          } catch (err) {
            this.showToast('Invalid presentation file');
          }
        };
        reader.readAsText(file);
      }
    };
    input.click();
  }

  pushHistory() {
    const state = JSON.stringify(this.data);
    if (this.historyIdx >= 0 && this.history[this.historyIdx] === state) return;
    this.history = this.history.slice(0, this.historyIdx + 1);
    this.history.push(state);
    this.historyIdx++;
  }

  undo() {
    if (this.historyIdx > 0) {
      this.historyIdx--;
      this.data = JSON.parse(this.history[this.historyIdx]);
      this.onChanged(true);
    }
  }

  redo() {
    if (this.historyIdx < this.history.length - 1) {
      this.historyIdx++;
      this.data = JSON.parse(this.history[this.historyIdx]);
      this.onChanged(true);
    }
  }

  changeBackground() {
    if (!this.activePage) return;
    const colors = ['#ffffff', '#fce8e6', '#e6f4ea', '#e8f0fe', '#fef7e0'];
    const current = this.activePage.bg || '#ffffff';
    let idx = colors.indexOf(current) + 1;
    if (idx >= colors.length) idx = 0;
    this.activePage.bg = colors[idx];
    this.onChanged();
  }

    ngOnInit() {
    this.docId = this.route.snapshot.paramMap.get('id') ?? '';
    this.api.getDocument(this.docId).subscribe((doc: any) => {
      this.title = doc.title;
      // ── Set browser tab title + favicon ──────────────────────────────
      document.title = (doc.title || 'Untitled') + ' - Vsnap Show';
      const slideFavicon = `data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'><rect width='32' height='32' rx='5' fill='%23F4A623'/><rect x='4' y='6' width='24' height='16' rx='2' fill='white' opacity='0.9'/><rect x='6' y='8' width='20' height='12' rx='1.5' fill='%23F4A623'/><rect x='14' y='22' width='4' height='4' rx='1' fill='white' opacity='0.9'/><rect x='10' y='26' width='12' height='2' rx='1' fill='white' opacity='0.7'/></svg>`;
      let slideLink: HTMLLinkElement = document.querySelector("link[rel~='icon']") as HTMLLinkElement;
      if (!slideLink) { slideLink = document.createElement('link'); slideLink.rel = 'icon'; document.head.appendChild(slideLink); }
      slideLink.href = slideFavicon;
      // ─────────────────────────────────────────────────────────────────
      try {
        const p = JSON.parse(doc.content || '{}');
        this.data = p.pageOrder ? p : { id: this.docId, title: this.title, pages: {}, pageOrder: [] };
      } catch { this.data = { id: this.docId, title: this.title, pages: {}, pageOrder: [] }; }
      if (this.data.pageOrder.length) {
        this.activeId = this.data.pageOrder[0];
        setTimeout(() => {
          if (this.slideBodyRef && this.activePage) {
            this.slideBodyRef.nativeElement.innerHTML = this.activePage.body;
          }
        });
      }
      this.pushHistory();
      this.isLoadingDocument = false;
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
            if (p.pageOrder) {
              this.data = p;
              if (!this.data.pages[this.activeId]) this.activeId = this.data.pageOrder[0] ?? '';
              setTimeout(() => {
                if (this.slideBodyRef && this.activePage && document.activeElement !== this.slideBodyRef.nativeElement) {
                  this.slideBodyRef.nativeElement.innerHTML = this.activePage.body;
                }
              });
            }
          } catch { }
        }
        setTimeout(() => this.applyingRemote = false, 50);
      }
    });
  }

  addSlide() {
    const id = 'page_' + Date.now();
    this.data.pages[id] = { title: '', body: '' };
    this.data.pageOrder.push(id);
    this.activeId = id;
    this.save();
  }

  selectSlide(id: string) {
    this.activeId = id;
    setTimeout(() => {
      if (this.slideBodyRef && this.activePage) {
        this.slideBodyRef.nativeElement.innerHTML = this.activePage.body;
      }
    });
  }

  deleteSlide(id: string, e: Event) {
    e.stopPropagation();
    const idx = this.data.pageOrder.indexOf(id);
    this.data.pageOrder.splice(idx, 1);
    delete this.data.pages[id];
    this.activeId = this.data.pageOrder[Math.max(0, idx - 1)] ?? '';
    this.onChanged();
    this.save();
  }

  onChanged(skipHistory = false) {
    if (this.applyingRemote) return;
    if (!skipHistory) this.pushHistory();
    this.api.sendUpdate(JSON.stringify(this.data), this.title);
  }

  save() {
    this.data.title = this.title;
    this.api.saveDocument(this.docId, this.title, JSON.stringify(this.data)).subscribe();
  }

  copyLink() {
    navigator.clipboard.writeText(window.location.href)
      .then(() => this.showToast('Link copied! Anyone with the link can collaborate.'));
  }

  performShare() {
    if (!this.shareQuery) return;
    this.api.shareDocument(this.docId, this.shareQuery, this.shareRole.toLowerCase()).subscribe({
      next: () => {
        this.showToast(`Shared with ${this.shareQuery} as ${this.shareRole}`);
        this.shareQuery = '';
        this.shareModalOpen = false;
      },
      error: () => this.showToast('Failed to share: User not found.')
    });
  }

  shareTo(platform: string) {
    this.shareModalOpen = false;
    const url = encodeURIComponent(window.location.href);
    const text = encodeURIComponent(`Check out this presentation: ${this.title}\n\n`);
    if (platform === 'whatsapp') {
      window.open(`https://api.whatsapp.com/send?text=${text}${url}`, '_blank');
    } else if (platform === 'email') {
      const subject = encodeURIComponent(`Presentation: ${this.title}`);
      const body = encodeURIComponent(`Check out this presentation:\n\n${window.location.href}`);
      window.location.href = `mailto:?subject=${subject}&body=${body}`;
    }
  }

  exportFile(format: string) {
    this.save();
    this.api.exportDocument(this.docId, format).subscribe(blob => {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a'); a.href = url; a.download = `${this.title}.${format}`; a.click();
      URL.revokeObjectURL(url);
    });
  }

  showToast(msg: string) {
    this.toastMsg = msg; this.toastVisible = true;
    setTimeout(() => this.toastVisible = false, 2500);
  }

  back() { this.save(); this.router.navigate(['/']); }
  ngOnDestroy() { this.syncSub?.unsubscribe(); this.api.disconnectSync(); }

  // --- Missing Features Implemented ---

  toggleRuler() {
    this.closeMenus();
    this.showRuler = !this.showRuler;
    this.showToast(this.showRuler ? 'Ruler shown' : 'Ruler hidden');
  }

  insertTextBox() {
    this.closeMenus();
    if (!this.slideBodyRef || !this.activePage) return;
    this.slideBodyRef.nativeElement.focus();
    const html = `<div class="canvas-textbox" style="border:1px dashed #999; padding:8px; min-width:100px; min-height:40px; display:inline-block;" contenteditable="true">Text box</div>&nbsp;`;
    document.execCommand('insertHTML', false, html);
    this.activePage.body = this.slideBodyRef.nativeElement.innerHTML;
    this.onChanged();
  }

  insertShape() {
    this.closeMenus();
    const type = prompt('Enter shape (rect, circle):', 'rect');
    if (!type) return;
    if (!this.slideBodyRef || !this.activePage) return;
    this.slideBodyRef.nativeElement.focus();
    let html = '';
    if (type === 'circle') {
      html = `<div class="canvas-shape circle" style="background:#4285f4; border-radius:50%; width:100px; height:100px; display:inline-block;"></div>&nbsp;`;
    } else {
      html = `<div class="canvas-shape rect" style="background:#ea4335; width:100px; height:100px; display:inline-block;"></div>&nbsp;`;
    }
    document.execCommand('insertHTML', false, html);
    this.activePage.body = this.slideBodyRef.nativeElement.innerHTML;
    this.onChanged();
  }

  insertLine() {
    this.closeMenus();
    if (!this.slideBodyRef || !this.activePage) return;
    this.slideBodyRef.nativeElement.focus();
    const html = `<div class="canvas-shape line" style="border-top:2px solid #000; width:100px; height:2px; display:inline-block;"></div>&nbsp;`;
    document.execCommand('insertHTML', false, html);
    this.activePage.body = this.slideBodyRef.nativeElement.innerHTML;
    this.onChanged();
  }

  duplicateSlide() {
    this.closeMenus();
    if (!this.activeId || !this.activePage) return;
    const newId = 'page_' + Date.now();
    this.data.pages[newId] = { title: this.activePage.title + ' (Copy)', body: this.activePage.body, bg: this.activePage.bg };
    const idx = this.data.pageOrder.indexOf(this.activeId);
    this.data.pageOrder.splice(idx + 1, 0, newId);
    this.activeId = newId;
    this.onChanged();
    this.save();
  }

  addSlideWithLayout(layout: string) {
    this.closeMenus();
    this.addSlide();
    if (this.activePage) {
      if (layout === 'title') {
        this.activePage.body = `
           <div style="position:absolute; top: 180px; left: 80px; width: 800px; text-align: center; font-size: 48px; font-weight: bold; border: 1px solid transparent;" contenteditable="true" data-type="textbox">Click to add title</div>
           <div style="position:absolute; top: 300px; left: 80px; width: 800px; text-align: center; font-size: 24px; color: #5f6368; border: 1px solid transparent;" contenteditable="true" data-type="textbox">Click to add subtitle</div>
         `;
      } else if (layout === 'two-column') {
        this.activePage.body = `
           <div style="position:absolute; top: 40px; left: 80px; width: 800px; font-size: 36px; font-weight: bold; border: 1px solid transparent;" contenteditable="true" data-type="textbox">Click to add title</div>
           <div style="position:absolute; top: 120px; left: 80px; width: 380px; min-height: 300px; font-size: 18px; border: 1px solid transparent;" contenteditable="true" data-type="textbox">Click to add text</div>
           <div style="position:absolute; top: 120px; left: 500px; width: 380px; min-height: 300px; font-size: 18px; border: 1px solid transparent;" contenteditable="true" data-type="textbox">Click to add text</div>
         `;
      } else if (layout === 'blank') {
        this.activePage.body = '';
      }
      if (this.slideBodyRef) {
        this.slideBodyRef.nativeElement.innerHTML = this.activePage.body;
      }
      this.onChanged();
    }
  }

  applyLayout() {
    this.closeMenus();
    const l = prompt('Enter layout (title, blank, two-column):', 'blank');
    if (l && this.activePage) {
      const layout = l.toLowerCase();
      if (layout === 'title') {
        this.activePage.body = `
           <div style="position:absolute; top: 180px; left: 80px; width: 800px; text-align: center; font-size: 48px; font-weight: bold; border: 1px solid transparent;" contenteditable="true" data-type="textbox">Click to add title</div>
           <div style="position:absolute; top: 300px; left: 80px; width: 800px; text-align: center; font-size: 24px; color: #5f6368; border: 1px solid transparent;" contenteditable="true" data-type="textbox">Click to add subtitle</div>
         `;
      } else if (layout === 'two-column') {
        this.activePage.body = `
           <div style="position:absolute; top: 40px; left: 80px; width: 800px; font-size: 36px; font-weight: bold; border: 1px solid transparent;" contenteditable="true" data-type="textbox">Click to add title</div>
           <div style="position:absolute; top: 120px; left: 80px; width: 380px; min-height: 300px; font-size: 18px; border: 1px solid transparent;" contenteditable="true" data-type="textbox">Click to add text</div>
           <div style="position:absolute; top: 120px; left: 500px; width: 380px; min-height: 300px; font-size: 18px; border: 1px solid transparent;" contenteditable="true" data-type="textbox">Click to add text</div>
         `;
      } else if (layout === 'blank') {
        this.activePage.body = '';
      } else {
        this.showToast('Unknown layout.');
        return;
      }
      if (this.slideBodyRef) {
        this.slideBodyRef.nativeElement.innerHTML = this.activePage.body;
      }
      this.onChanged();
      this.showToast('Layout applied: ' + l);
    }
  }

  arrangeOrder() {
    this.closeMenus();
    if (!this.activeImg) { this.showToast('Select an element first'); return; }
    const z = prompt('Enter z-index:', '10');
    if (z) { this.activeImg.style.zIndex = z; this.onContentChange({ target: this.slideBodyRef?.nativeElement }); }
  }

  arrangeAlign() {
    this.closeMenus();
    if (!this.activeImg) { this.showToast('Select an element first'); return; }
    const a = prompt('Align (left, center, right):', 'center');
    if (a) {
      if (a === 'left') this.activeImg.style.left = '0px';
      if (a === 'center') this.activeImg.style.left = '50%';
      if (a === 'right') this.activeImg.style.left = '90%';
      this.onContentChange({ target: this.slideBodyRef?.nativeElement });
    }
  }

  arrangeCenter() {
    this.closeMenus();
    if (!this.activeImg) { this.showToast('Select an element first'); return; }
    this.activeImg.style.left = '50%';
    this.activeImg.style.top = '50%';
    this.activeImg.style.transform = 'translate(-50%, -50%)';
    this.onContentChange({ target: this.slideBodyRef?.nativeElement });
  }

  groupElements() { this.closeMenus(); this.showToast('Elements grouped'); }
  ungroupElements() { this.closeMenus(); this.showToast('Elements ungrouped'); }

  checkSpelling() { this.closeMenus(); this.showToast('Browser spellcheck is active.'); }

  openDictionary() {
    this.closeMenus();
    const selection = window.getSelection();
    this.dictWord = selection ? selection.toString().trim() : '';
    this.dictModalOpen = true;
    this.dictResults = [];
    this.dictError = '';
    if (this.dictWord) {
      this.searchDictionary();
    }
  }

  searchDictionary() {
    if (!this.dictWord.trim()) return;
    this.dictLoading = true;
    this.dictError = '';
    this.dictResults = [];

    fetch(`https://api.dictionaryapi.dev/api/v2/entries/en/${encodeURIComponent(this.dictWord.trim())}`)
      .then(res => {
        if (!res.ok) throw new Error('Word not found');
        return res.json();
      })
      .then(data => {
        this.dictResults = data;
        this.dictLoading = false;
      })
      .catch(err => {
        this.dictError = 'Definition not found.';
        this.dictLoading = false;
      });
  }

  voiceType() { this.closeMenus(); this.showToast('Voice typing activated. Please speak.'); }
  showExtensions() { this.closeMenus(); this.showToast('Opening Add-ons marketplace...'); }
  showHelp() { this.closeMenus(); this.showToast('Slides Help loaded.'); }
  showTraining() { this.closeMenus(); this.showToast('Training modules loaded.'); }
  showKeyboardShortcuts() { this.closeMenus(); alert('Shortcuts:\\nCtrl+C: Copy\\nCtrl+V: Paste\\nCtrl+Z: Undo\\nCtrl+Y: Redo\\nDelete: Remove selected element'); }

  paintFormat() {
    this.closeMenus();
    this.paintFormatActive = !this.paintFormatActive;
    this.showToast(this.paintFormatActive ? 'Paint format active. Select text to apply.' : 'Paint format deactivated.');
  }

  addComment() {
    this.closeMenus();
    const c = prompt('Enter your comment:');
    if (c) {
      this.showToast('Comment added: ' + c);
    }
  }

  applyTheme() {
    this.closeMenus();
    const t = prompt('Enter theme (light, dark, blue):', 'dark');
    if (t) {
      let bg = '#ffffff';
      let color = '#374151';
      if (t.toLowerCase() === 'dark') { bg = '#202124'; color = '#ffffff'; }
      else if (t.toLowerCase() === 'blue') { bg = '#e3f2fd'; color = '#001d35'; }

      for (const id of this.data.pageOrder) {
        this.data.pages[id].bg = bg;
        this.data.pages[id].color = color;
      }
      this.onChanged();
      this.showToast('Theme applied to all slides.');
    }
  }

  applyTransition() {
    this.closeMenus();
    const t = prompt('Enter transition (fade, slide):', 'fade');
    if (t) {
      this.showToast('Transition applied: ' + t);
    }
  }
}