import { Component, OnInit } from '@angular/core';
import { Router } from '@angular/router';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpEventType } from '@angular/common/http';
import { ChatWidgetComponent } from '../../components/chat-widget/chat-widget.component';
import { ApiService } from '../../services/api.service';
import { AuthService } from '../../services/auth.service';
import { filter, take } from 'rxjs/operators';
@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [CommonModule, FormsModule, ChatWidgetComponent],
  template: `
    <div class="layout" (click)="closeAllMenus()">
      <!-- Top Header -->
      <header class="top-header">
        <div class="header-left">
          <div class="app-launcher" (click)="toggleMenu('launcher', $event)">
            <svg *ngIf="currentApp === 'Writer'" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="8" y1="13" x2="16" y2="13"></line><line x1="8" y1="17" x2="16" y2="17"></line></svg>
            <svg *ngIf="currentApp === 'Sheet'" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#059669" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M4 20v-5a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v5Z"></path><path d="M10 20V9a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v11Z"></path><path d="M16 20V5a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v15Z"></path></svg>
            <svg *ngIf="currentApp === 'Show'" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#d97706" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="12" rx="2" ry="2"></rect><path d="M8 22l4-6 4 6M12 16v-4"></path><path d="M7 10h10M12 10v3"></path></svg>
            <svg *ngIf="currentApp === 'WorkDrive'" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#6366f1" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"></path></svg>
            <span class="app-name">{{ currentApp }}</span>
            <span class="material-symbols-outlined" style="font-size:16px; color:#5f6368; margin-left:4px;">arrow_drop_down</span>
            
            <!-- Launcher Dropdown -->
            <div class="launcher-dropdown shadow-lg" *ngIf="menus['launcher']" (click)="$event.stopPropagation()">
              <div class="ld-item" [class.active]="currentApp === 'Writer'" (click)="setApp('Writer', $event)">
                 <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="8" y1="13" x2="16" y2="13"></line><line x1="8" y1="17" x2="16" y2="17"></line></svg>
                 <div class="ld-text">
                   <div class="ld-title">Writer</div>
                   <div class="ld-desc">A clean, crisp space to write.</div>
                 </div>
              </div>
              <div class="ld-item" [class.active]="currentApp === 'Sheet'" (click)="setApp('Sheet', $event)">
                 <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#059669" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M4 20v-5a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v5Z"></path><path d="M10 20V9a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v11Z"></path><path d="M16 20V5a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v15Z"></path></svg>
                 <div class="ld-text">
                   <div class="ld-title">Sheet</div>
                   <div class="ld-desc">The spreadsheet application.</div>
                 </div>
              </div>
              <div class="ld-item" [class.active]="currentApp === 'Show'" (click)="setApp('Show', $event)">
                 <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#d97706" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="12" rx="2" ry="2"></rect><path d="M8 22l4-6 4 6M12 16v-4"></path><path d="M7 10h10M12 10v3"></path></svg>
                 <div class="ld-text">
                   <div class="ld-title">Show</div>
                   <div class="ld-desc">Broadcast presentations.</div>
                 </div>
              </div>
              <div class="ld-item" [class.active]="currentApp === 'WorkDrive'" (click)="setApp('WorkDrive', $event)">
                 <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#6366f1" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"></path></svg>
                 <div class="ld-text">
                   <div class="ld-title">WorkDrive</div>
                   <div class="ld-desc">Share and manage files securely.</div>
                 </div>
              </div>
            </div>
          </div>
        </div>

        <div class="header-center">
          <div class="search-bar" style="position:relative;">
            <span class="material-symbols-outlined search-icon">search</span>
            <input type="text" placeholder="Search documents" [(ngModel)]="searchQuery" (focus)="searchFocused=true" (blur)="onSearchBlur()">
            
            <div class="search-dropdown shadow-lg" *ngIf="searchFocused && searchResults.length > 0">
              <div class="search-item" *ngFor="let res of searchResults" (mousedown)="$event.preventDefault(); open(res)">
                <svg *ngIf="res.doc_type === 'sheet'" class="search-res-icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#059669" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M4 20v-5a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v5Z"></path><path d="M10 20V9a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v11Z"></path><path d="M16 20V5a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v15Z"></path></svg>
                <svg *ngIf="res.doc_type === 'doc'" class="search-res-icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="8" y1="13" x2="16" y2="13"></line><line x1="8" y1="17" x2="16" y2="17"></line></svg>
                <svg *ngIf="res.doc_type === 'slide'" class="search-res-icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#d97706" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="12" rx="2" ry="2"></rect><path d="M8 22l4-6 4 6M12 16v-4"></path><path d="M7 10h10M12 10v3"></path></svg>
                <div class="search-res-text">
                  <div class="srt-title">{{ res.title }}</div>
                  <div class="srt-owner">{{ isOwner(res) ? 'me' : (res.owner_name || 'Unknown') }}</div>
                </div>
                <div class="search-res-date">Last modified on {{ formatDate(res.updated_at) }}</div>
              </div>
              <div class="search-footer">
                 <span style="color:#0f9d58; cursor:pointer;">Search across zoho <span class="material-symbols-outlined" style="font-size:14px;vertical-align:middle;">open_in_new</span></span>
              </div>
            </div>
          </div>
        </div>

        <div class="header-right">
          <div style="position:relative;">
            <input type="file" #fileInput (change)="onUpload($event)" style="display:none">
            <button class="btn btn-outline-primary" (click)="toggleMenu('upload', $event)">
              <span class="material-symbols-outlined" style="font-size:18px;">upload</span> Upload
              <span class="material-symbols-outlined" style="font-size:16px;">arrow_drop_down</span>
            </button>
            <div class="dropdown-panel upload-panel shadow-lg" *ngIf="menus['upload']" (click)="$event.stopPropagation()">
              <div class="up-item" (click)="fileInput.click(); menus['upload']=false">Upload File</div>
              <div class="up-item" (click)="showToast('Folder upload not supported yet'); menus['upload']=false">Upload Folder</div>
            </div>
          </div>
          
          <div style="position:relative;">
            <button class="btn btn-primary" (click)="toggleMenu('newMenu', $event)">
              <span class="material-symbols-outlined" style="font-size:18px;">add</span> Create New
            </button>
            <div class="new-menu shadow-lg" *ngIf="menus['newMenu']" (click)="$event.stopPropagation()">
              <a (click)="create('slide')"><span class="material-symbols-outlined" style="color:#f59e0b;">slideshow</span> Presentation</a>
            </div>
          </div>
          
          <!-- Notifications -->
          <div style="position:relative;">
            <span class="material-symbols-outlined icon-action" (click)="toggleMenu('notif', $event)">notifications</span>
            <div class="dropdown-panel notif-panel shadow-lg" *ngIf="menus['notif']" (click)="$event.stopPropagation()">
              <div class="dp-header">NOTIFICATION</div>
              <div class="dp-body empty-notif">
                You haven't received any notifications yet
              </div>
              <div class="dp-footer">
                <a href="javascript:void(0)" class="text-blue">Mark all as read</a>
              </div>
            </div>
          </div>

          <!-- Profile -->
          <div style="position:relative;">
            <div class="user-avatar" (click)="toggleMenu('profile', $event)">{{ initials }}</div>
            <div class="dropdown-panel profile-panel shadow-lg" *ngIf="menus['profile']" (click)="$event.stopPropagation()">
              <div class="prof-header">
                <div class="prof-avatar-lg">{{ initials }}</div>
                <div class="prof-email">{{ auth.user?.email || 'admin@vsnapmail.co.in' }}</div>
                <div class="prof-id">User ID: {{ auth.user?.id || 1 }}</div>
              </div>
              <div class="prof-body">
                <div class="prof-item">
                  <div style="display:flex; align-items:center; gap:12px;">
                    <span class="material-symbols-outlined">person</span> My account
                  </div>
                  <span class="prof-action-txt" (click)="auth.logout()">Sign Out</span>
                </div>
                <div class="prof-item"><span class="material-symbols-outlined">desktop_windows</span> Admin panel</div>
                <div class="prof-item"><span class="material-symbols-outlined">accessibility</span> Accessibility options</div>
                <div class="prof-item active-prof"><span class="material-symbols-outlined">code</span> Deluge Functions</div>
                <div class="prof-item"><span class="material-symbols-outlined">help</span> User Guide</div>
                <div class="prof-item"><span class="material-symbols-outlined">lightbulb</span> Knowledge Base</div>
                <div class="prof-item"><span class="material-symbols-outlined">feedback</span> Feedback</div>
                <div class="prof-item"><span class="material-symbols-outlined">forum</span> Community</div>
                <div class="prof-item"><span class="material-symbols-outlined">edit</span> Blogs</div>
                <div class="prof-item" style="justify-content:space-between;">
                   <div style="display:flex; align-items:center; gap:12px;">
                     <span class="material-symbols-outlined">warning</span> Featured Articles
                   </div>
                   <span class="material-symbols-outlined" style="font-size:18px;">chevron_right</span>
                </div>
              </div>
            </div>
          </div>

          <!-- Apps Grid -->
          <div style="position:relative;">
            <span class="material-symbols-outlined icon-action" (click)="toggleMenu('apps', $event)">apps</span>
            <div class="dropdown-panel apps-panel shadow-lg" *ngIf="menus['apps']" (click)="$event.stopPropagation()">
              <div class="ld-search">
                 <span class="material-symbols-outlined">search</span>
                 <input type="text" placeholder="Search Applications">
              </div>
              <div class="ld-section">ALL ZOHO APPS</div>
              <div class="ld-subsection">EMAIL & COLLABORATION</div>
              <div class="ld-grid">
                 <div class="ld-icon-item" (click)="toggleWidget('calendar'); menus['apps']=false">
                    <span class="material-symbols-outlined" style="color:#d32f2f;">calendar_month</span>
                    <span>Calendar</span>
                 </div>
                 <div class="ld-icon-item" (click)="toggleWidget('notes'); menus['apps']=false">
                    <span class="material-symbols-outlined" style="color:#1976d2;">notes</span>
                    <span>Notes</span>
                 </div>
                 <div class="ld-icon-item" (click)="toggleWidget('scratchpad'); menus['apps']=false">
                    <span class="material-symbols-outlined" style="color:#fbc02d;">edit_note</span>
                    <span>Scratchpad</span>
                 </div>
                 <div class="ld-icon-item" (click)="toggleWidget('chat'); menus['apps']=false">
                    <span class="material-symbols-outlined" style="color:#1976d2;">chat</span>
                    <span>Chat</span>
                 </div>
              </div>
            </div>
          </div>
        </div>
      </header>

      <!-- Main Content -->
      <main class="main-content">
        
        <!-- Navigation Tabs -->
        <div class="nav-toolbar">
          <div class="nav-tabs">
            <a class="nav-tab" [class.active]="currentTab === 'Recents'" (click)="currentTab = 'Recents'">Recents</a>
            <a class="nav-tab" [class.active]="currentTab === 'My Documents'" (click)="currentTab = 'My Documents'">My Documents</a>
            <a class="nav-tab" [class.active]="currentTab === 'Shared with me'" (click)="currentTab = 'Shared with me'">Shared with me</a>
            <a class="nav-tab" [class.active]="currentTab === 'Favorites'" (click)="currentTab = 'Favorites'">Favorites</a>
            <a class="nav-tab" [class.active]="currentTab === 'Trash'" (click)="currentTab = 'Trash'">Trash</a>
          </div>
          <div class="view-controls">
            <div class="vc-btn" [class.active]="viewMode === 'grid'" (click)="viewMode = 'grid'"><span class="material-symbols-outlined">grid_view</span></div>
            <div class="vc-btn" [class.active]="viewMode === 'list'" (click)="viewMode = 'list'"><span class="material-symbols-outlined">view_list</span></div>
          </div>
        </div>

        <!-- Document Views -->
        <div *ngIf="filteredDocs.length > 0; else empty">
          
          <!-- Grid View -->
          <div class="grid-container" *ngIf="viewMode === 'grid'">
            <div class="grid-section-title">Last opened earlier this week</div>
            <div class="grid-layout">
              <div class="grid-card" *ngFor="let doc of filteredDocs; trackBy: trackById" (click)="open(doc)">
                <div class="gc-header">
                  <span class="material-symbols-outlined gc-star" [class.is-fav]="doc._favorite" (click)="toggleFavorite(doc, $event)">
                    {{ doc._favorite ? 'star' : 'star_outline' }}
                  </span>
                  <svg *ngIf="doc.doc_type === 'sheet'" class="doc-icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#059669" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M4 20v-5a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v5Z"></path><path d="M10 20V9a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v11Z"></path><path d="M16 20V5a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v15Z"></path></svg>
                  <svg *ngIf="doc.doc_type === 'doc'" class="doc-icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="8" y1="13" x2="16" y2="13"></line><line x1="8" y1="17" x2="16" y2="17"></line></svg>
                  <svg *ngIf="doc.doc_type === 'slide'" class="doc-icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#d97706" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="12" rx="2" ry="2"></rect><path d="M8 22l4-6 4 6M12 16v-4"></path><path d="M7 10h10M12 10v3"></path></svg>
                  <span class="gc-title">{{ doc.title }}</span>
                </div>
                <div class="gc-footer" style="display:flex; justify-content:space-between; align-items:center;">
                  <span class="gc-date">{{ formatDate(doc.updated_at) }}</span>
                  <span *ngIf="currentTab === 'Trash'" class="material-symbols-outlined" style="font-size:16px; cursor:pointer;" title="Restore" (click)="restore(doc, $event); $event.stopPropagation()">restore</span>
                </div>
              </div>
            </div>
          </div>

          <!-- List View -->
          <div class="list-container" *ngIf="viewMode === 'list'">
            <div class="list-header">
              <div class="col-name">Name</div>
              <div class="col-owner">Owner</div>
              <div class="col-date">Last Modified</div>
              <div class="col-actions"></div>
            </div>
            <div class="list-body">
              <div class="list-row" *ngFor="let doc of filteredDocs; trackBy: trackById" (click)="open(doc)">
                <div class="col-name">
                  <span class="material-symbols-outlined fav-icon" [class.is-fav]="doc._favorite" (click)="toggleFavorite(doc, $event)">
                    {{ doc._favorite ? 'star' : 'star_outline' }}
                  </span>
                  <svg *ngIf="doc.doc_type === 'sheet'" class="doc-icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#059669" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M4 20v-5a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v5Z"></path><path d="M10 20V9a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v11Z"></path><path d="M16 20V5a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v15Z"></path></svg>
                  <svg *ngIf="doc.doc_type === 'doc'" class="doc-icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="8" y1="13" x2="16" y2="13"></line><line x1="8" y1="17" x2="16" y2="17"></line></svg>
                  <svg *ngIf="doc.doc_type === 'slide'" class="doc-icon" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#d97706" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="12" rx="2" ry="2"></rect><path d="M8 22l4-6 4 6M12 16v-4"></path><path d="M7 10h10M12 10v3"></path></svg>
                  <span class="doc-title-text">{{ doc.title }}</span>
                </div>
                <div class="col-owner">
                  <div class="owner-avatar" [style.background]="isOwner(doc) ? '#1a73e8' : '#8ab4f8'">{{ isOwner(doc) ? 'M' : (doc.owner_name ? doc.owner_name.charAt(0).toUpperCase() : 'U') }}</div>
                  <span class="owner-name">{{ isOwner(doc) ? 'me' : (doc.owner_name || 'Unknown') }}</span>
                </div>
                <div class="col-date">
                  <span class="date-text">{{ formatDate(doc.updated_at) }}</span>
                </div>
                <div class="col-actions">
                  <button *ngIf="currentTab !== 'Trash'" class="action-btn" title="Copy Link" (click)="copyLink(doc); $event.stopPropagation()">
                    <span class="material-symbols-outlined">link</span>
                  </button>
                  <button *ngIf="currentTab === 'Trash'" class="action-btn" title="Restore" (click)="restore(doc, $event); $event.stopPropagation()">
                    <span class="material-symbols-outlined">restore</span>
                  </button>
                  <button class="action-btn" title="Delete" (click)="delete(doc); $event.stopPropagation()">
                    <span class="material-symbols-outlined">delete</span>
                  </button>
                </div>
              </div>
            </div>
          </div>

        </div>
        <ng-template #empty>
          <div class="empty-state">
            <p>No documents found.</p>
          </div>
        </ng-template>
      </main>

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

      <div class="widget-panel shadow-lg" *ngIf="activeWidget === 'calendar'">
         <div class="wp-header">
            <span>Calendar Filter</span>
            <span class="material-symbols-outlined" (click)="activeWidget=null" style="cursor:pointer;">close</span>
         </div>
         <div class="wp-body">
            <div class="cal-controls">
               <span class="material-symbols-outlined" (click)="changeMonth(-1)" style="cursor:pointer;">chevron_left</span>
               <span class="cal-month-yr">{{ currentMonthName }} {{ calYear }}</span>
               <span class="material-symbols-outlined" (click)="changeMonth(1)" style="cursor:pointer;">chevron_right</span>
            </div>
            <div class="cal-grid">
               <div class="cal-day-name">Su</div><div class="cal-day-name">Mo</div><div class="cal-day-name">Tu</div>
               <div class="cal-day-name">We</div><div class="cal-day-name">Th</div><div class="cal-day-name">Fr</div><div class="cal-day-name">Sa</div>
               <div class="cal-cell" *ngFor="let d of calendarDays" [class.empty]="!d" [class.selected]="d && isSameDate(d, filterDate)" [class.has-doc]="d && hasDocsOnDate(d)" (click)="d && selectDate(d)">
                  {{ d ? d.getDate() : '' }}
               </div>
            </div>
            <div style="text-align:center; margin-top:12px;" *ngIf="filterDate">
               <button class="btn btn-sm btn-outline-primary" (click)="filterDate=null">Clear Filter</button>
            </div>
         </div>
      </div>

      <div class="widget-panel shadow-lg" *ngIf="activeWidget === 'notes'">
         <div class="wp-header">
            <span>Notes</span>
            <span class="material-symbols-outlined" (click)="activeWidget=null" style="cursor:pointer;">close</span>
         </div>
         <div class="wp-body" style="padding:0; height:100%;">
            <textarea class="widget-textarea" placeholder="Take a quick note..." [(ngModel)]="notesContent"></textarea>
         </div>
      </div>

      <div class="widget-panel shadow-lg" *ngIf="activeWidget === 'scratchpad'">
         <div class="wp-header">
            <span>Scratchpad</span>
            <span class="material-symbols-outlined" (click)="activeWidget=null" style="cursor:pointer;">close</span>
         </div>
         <div class="wp-body" style="padding:0; height:100%;">
            <textarea class="widget-textarea" placeholder="Scribble here..." [(ngModel)]="scratchpadContent" style="font-family: monospace;"></textarea>
         </div>
      </div>

      <div class="toast" [class.show]="toastVisible">{{ toastMsg }}</div>

      <!-- Delete Confirmation Modal -->
      <div class="modal-backdrop" *ngIf="deleteConfirmDoc" (click)="cancelDelete()">
        <div class="delete-modal shadow-lg" (click)="$event.stopPropagation()">
          <div class="dm-header">
            <span class="material-symbols-outlined" style="color: #d32f2f;">warning</span>
            <span>Confirm Deletion</span>
          </div>
          <div class="dm-body">
            <p *ngIf="deleteConfirmPermanent">Are you sure you want to permanently delete <strong>{{ deleteConfirmDoc.title }}</strong>? This action cannot be undone.</p>
            <p *ngIf="!deleteConfirmPermanent">Are you sure you want to move <strong>{{ deleteConfirmDoc.title }}</strong> to the trash?</p>
          </div>
          <div class="dm-footer">
            <button class="btn btn-outline-primary" (click)="cancelDelete()">Cancel</button>
            <button class="btn btn-danger" (click)="confirmDelete()">{{ deleteConfirmPermanent ? 'Permanently Delete' : 'Move to Trash' }}</button>
          </div>
        </div>
      </div>

      <!-- Upload Progress Overlay (2026 Premium Style) -->
      <div class="upload-overlay" *ngIf="isUploading">
        <div class="upload-modal shadow-lg">
           <div class="um-icon">
              <span class="material-symbols-outlined" style="font-size:32px; color:#1a73e8;">cloud_upload</span>
           </div>
           <div class="um-title">Uploading Document...</div>
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
    </div>
  `,
  styles: [`
    :host { display: block; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; }
    .layout { display: flex; flex-direction: column; min-height: 100vh; background: #ffffff; color: #202124; padding-bottom: 36px; }
    
    /* Top Header */
    .top-header { height: 60px; background: #ffffff; border-bottom: 1px solid #e0e0e0; display: flex; align-items: center; justify-content: space-between; padding: 0 24px; position: relative; z-index: 50; gap: 20px; }
    
    .header-left { display: flex; align-items: center; flex-shrink: 0; }
    .header-center { flex: 1; display: flex; justify-content: center; max-width: 600px; }
    .header-right { display: flex; align-items: center; justify-content: flex-end; gap: 16px; flex-shrink: 0; }

    /* App Launcher */
    .app-launcher { display: flex; align-items: center; gap: 8px; cursor: pointer; padding: 6px 12px; border-radius: 4px; transition: background 0.2s; position: relative; }
    .app-launcher:hover { background: #f1f3f4; }
    .app-name { font-size: 18px; font-weight: 500; color: #202124; letter-spacing: -0.5px; }
    
    .launcher-dropdown { position: absolute; top: 100%; left: 0; background: #fff; width: 320px; border-radius: 8px; border: 1px solid #e0e0e0; z-index: 100; padding: 8px 0; overflow: hidden; display: flex; flex-direction: column; }
    .ld-item { display: flex; align-items: flex-start; gap: 16px; padding: 12px 24px; cursor: pointer; transition: background 0.2s; }
    .ld-item:hover { background: #f8f9fa; }
    .ld-item.active { background: #e8f0fe; }
    .ld-item .material-symbols-outlined { font-size: 24px; margin-top: 2px; }
    .ld-title { font-size: 14px; font-weight: 500; color: #202124; margin-bottom: 2px; }
    .ld-desc { font-size: 12px; color: #5f6368; line-height: 1.4; }
    
    .apps-panel { width: 340px; padding: 0; background: #fff; border-radius: 8px; overflow: hidden; display: flex; flex-direction: column; }
    .ld-search { padding: 12px 16px; position: relative; }
    .ld-search input { width: 100%; border: 1px solid #dadce0; border-radius: 20px; padding: 8px 12px 8px 36px; outline: none; font-size: 13px; transition: box-shadow 0.2s; }
    .ld-search input:focus { box-shadow: 0 1px 3px rgba(0,0,0,0.1); border-color: #1a73e8; }
    .ld-search .material-symbols-outlined { position: absolute; left: 24px; top: 18px; color: #5f6368; font-size: 18px; }
    .ld-section { font-size: 11px; font-weight: 600; color: #202124; padding: 8px 16px; letter-spacing: 0.5px; }
    .ld-subsection { font-size: 10px; font-weight: 600; color: #5f6368; padding: 8px 16px; text-transform: uppercase; }
    .ld-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; padding: 0 16px 16px 16px; }
    .ld-icon-item { display: flex; align-items: center; gap: 8px; padding: 8px 12px; cursor: pointer; border-radius: 6px; transition: background 0.2s; border: 1px solid transparent; }
    .ld-icon-item:hover { background: #f8f9fa; border-color: #dadce0; }
    .ld-icon-item span:last-child { font-size: 13px; color: #202124; font-weight: 500; }
    .ld-icon-item.active { background: #e8f0fe; border-color: #d2e3fc; }
    
    /* Bottom Chat Bar */
    .bottom-chat-bar { position: fixed; bottom: 0; left: 0; right: 0; display: flex; background: #f8f9fa; border-top: 1px solid #e0e0e0; z-index: 1000; height: 36px; box-shadow: 0 -1px 3px rgba(0,0,0,0.05); }
    .bcb-item { display: flex; align-items: center; gap: 8px; padding: 0 16px; cursor: pointer; border-right: 1px solid #e0e0e0; font-size: 13px; font-weight: 500; color: #202124; transition: background 0.2s; position: relative; }
    .bcb-item:hover { background: #e8f0fe; }
    .bcb-item .material-symbols-outlined { font-size: 18px; }
    .bcb-badge { position: absolute; top: -6px; left: 16px; background: #d32f2f; color: #fff; font-size: 10px; font-weight: bold; border-radius: 50%; width: 16px; height: 16px; display: flex; align-items: center; justify-content: center; }

    /* Widgets */
    .widget-panel { position: fixed; bottom: 48px; right: 24px; width: 300px; background: #fff; border-radius: 8px; border: 1px solid #e0e0e0; z-index: 1100; display: flex; flex-direction: column; overflow: hidden; box-shadow: 0 4px 12px rgba(0,0,0,0.15); }
    .wp-header { display: flex; justify-content: space-between; align-items: center; padding: 12px 16px; border-bottom: 1px solid #e0e0e0; font-weight: 500; color: #202124; background: #f8f9fa; }
    .wp-body { padding: 16px; flex: 1; min-height: 200px; display: flex; flex-direction: column; }
    .widget-textarea { flex: 1; border: none; outline: none; padding: 16px; resize: none; width: 100%; height: 200px; color: #202124; }
    .cal-controls { display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; font-weight: 500; }
    .cal-month-yr { font-size: 14px; }
    .cal-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 4px; text-align: center; }
    .cal-day-name { font-size: 11px; color: #5f6368; font-weight: 600; margin-bottom: 4px; }
    .cal-cell { font-size: 13px; padding: 6px 0; border-radius: 50%; cursor: pointer; color: #202124; transition: background 0.2s; position: relative; }
    .cal-cell:not(.empty):hover { background: #f1f3f4; }
    .cal-cell.selected { background: #1a73e8; color: #fff; font-weight: 500; }
    .cal-cell.has-doc::after { content: ''; position: absolute; bottom: 2px; left: 50%; transform: translateX(-50%); width: 4px; height: 4px; background: #10b981; border-radius: 50%; }
    
    /* Search */
    .search-bar { width: 100%; max-width: 500px; display: flex; align-items: center; background: #ffffff; border: 1px solid #dadce0; border-radius: 24px; padding: 0 16px; height: 44px; transition: background 0.2s, box-shadow 0.2s, border-color 0.2s; position: relative; }
    .search-bar:focus-within { box-shadow: 0 1px 3px rgba(0,0,0,0.1); border-color: #1a73e8; }
    .search-icon { color: #5f6368; margin-right: 8px; font-size: 20px; }
    .search-bar input { flex: 1; border: none; background: transparent; outline: none; font-size: 14px; color: #202124; }
    
    .btn { display: flex; align-items: center; gap: 6px; padding: 0 16px; height: 36px; border-radius: 4px; font-size: 14px; font-weight: 500; cursor: pointer; transition: all 0.2s; }
    .btn-outline-primary { background: #ffffff; border: 1px solid #1a73e8; color: #1a73e8; }
    .btn-outline-primary:hover { background: rgba(26,115,232,0.04); }
    .btn-primary { background: #1a73e8; border: none; color: #ffffff; }
    .btn-primary:hover { background: #1557b0; box-shadow: 0 1px 3px rgba(0,0,0,0.2); }
    
    .new-menu { position: absolute; top: 44px; right: 0; width: 200px; background: #fff; border-radius: 4px; padding: 8px 0; border: 1px solid #e0e0e0; z-index: 100; }
    .new-menu a { display: flex; align-items: center; gap: 12px; padding: 10px 16px; cursor: pointer; color: #202124; font-size: 14px; transition: background 0.2s; }
    .new-menu a:hover { background: #f1f3f4; }
    
    .icon-action { color: #5f6368; font-size: 22px; cursor: pointer; padding: 6px; border-radius: 50%; transition: background 0.2s; }
    .icon-action:hover { background: #f1f3f4; }
    
    .user-avatar { width: 32px; height: 32px; border-radius: 50%; background: #1a73e8; color: #fff; display: flex; align-items: center; justify-content: center; font-weight: 600; font-size: 13px; cursor: pointer; margin: 0 8px; transition: box-shadow 0.2s; }
    .user-avatar:hover { box-shadow: 0 1px 3px rgba(0,0,0,0.3); }

    /* Dropdown Panels */
    .dropdown-panel { position: absolute; top: 44px; right: 0; background: #fff; border-radius: 8px; border: 1px solid #e0e0e0; z-index: 100; overflow: hidden; display: flex; flex-direction: column; }
    
    /* Notifications */
    .notif-panel { width: 320px; }
    .dp-header { font-size: 11px; font-weight: 600; color: #202124; padding: 12px 16px; border-bottom: 1px solid #e0e0e0; letter-spacing: 0.5px; }
    .dp-body { padding: 48px 24px; text-align: center; color: #5f6368; font-size: 13px; line-height: 1.5; }
    .dp-footer { border-top: 1px solid #e0e0e0; padding: 12px; text-align: center; }
    .text-blue { color: #1a73e8; font-size: 13px; text-decoration: none; font-weight: 500; }
    .text-blue:hover { text-decoration: underline; }
    
    /* Profile */
    .profile-panel { width: 300px; max-height: calc(100vh - 80px); overflow-y: auto; }
    .prof-header { display: flex; flex-direction: column; align-items: center; padding: 24px 16px 16px; border-bottom: 1px solid #e0e0e0; }
    .prof-avatar-lg { width: 64px; height: 64px; border-radius: 50%; background: #1a73e8; color: #fff; display: flex; align-items: center; justify-content: center; font-size: 28px; font-weight: 500; margin-bottom: 12px; }
    .prof-email { font-size: 14px; color: #202124; margin-bottom: 4px; }
    .prof-id { font-size: 12px; color: #5f6368; }
    .prof-body { padding: 8px 0; }
    .prof-item { display: flex; align-items: center; gap: 12px; padding: 10px 24px; font-size: 14px; color: #3c4043; cursor: pointer; transition: background 0.2s; }
    .prof-item:hover { background: #f1f3f4; }
    .active-prof { background: #f8f9fa; }
    .prof-item .material-symbols-outlined { color: #5f6368; font-size: 20px; }
    .prof-action-txt { color: #d93025; font-weight: 500; font-size: 13px; }
    .prof-action-txt:hover { text-decoration: underline; }

    /* Apps Grid */
    .apps-panel { width: 320px; padding: 16px 0; }
    .apps-search-wrapper { padding: 0 16px 16px; }
    .apps-search { display: flex; align-items: center; border: 1px solid #dadce0; border-radius: 20px; padding: 0 12px; height: 36px; transition: border-color 0.2s; }
    .apps-search:focus-within { border-color: #1a73e8; }
    .apps-search .material-symbols-outlined { color: #5f6368; font-size: 18px; margin-right: 8px; }
    .apps-search input { flex: 1; border: none; outline: none; font-size: 13px; color: #202124; }
    .apps-section { padding: 0 16px; }
    .apps-title { font-size: 11px; font-weight: 600; color: #202124; letter-spacing: 0.5px; margin-bottom: 12px; }
    .mt-2 { margin-top: 16px; }
    .apps-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
    .app-item { display: flex; align-items: center; gap: 10px; font-size: 13px; color: #202124; cursor: pointer; padding: 4px; border-radius: 4px; transition: background 0.2s; }
    .app-item:hover { background: #f1f3f4; }
    .app-item .material-symbols-outlined { font-size: 22px; }

    /* Search Dropdown */
    .search-dropdown { position: absolute; top: calc(100% + 4px); left: 0; right: 0; background: #fff; border-radius: 12px; border: 1px solid #10b981; z-index: 100; overflow: hidden; display: flex; flex-direction: column; }
    .search-item { display: flex; align-items: center; padding: 12px 16px; cursor: pointer; transition: background 0.2s; }
    .search-item:hover { background: #f8f9fa; }
    .search-res-icon { font-size: 20px; margin-right: 12px; }
    .search-res-text { flex: 1; }
    .srt-title { font-size: 13px; font-weight: 500; color: #202124; }
    .srt-owner { font-size: 11px; color: #5f6368; }
    .search-res-date { font-size: 12px; color: #5f6368; }
    .search-footer { border-top: 1px solid #e0e0e0; padding: 12px 16px; text-align: right; font-size: 13px; font-weight: 500; background: #f8f9fa; }
    
    /* Upload Menu */
    .upload-panel { width: 160px; padding: 8px 0; }
    .up-item { padding: 10px 16px; font-size: 14px; color: #202124; cursor: pointer; transition: background 0.2s; }
    .up-item:hover { background: #f1f3f4; }
    
    /* Main Content */
    .main-content { flex: 1; display: flex; flex-direction: column; padding: 0 64px; }
    
    /* Nav Toolbar */
    .nav-toolbar { display: flex; align-items: center; justify-content: space-between; margin-top: 32px; border-bottom: 1px solid #e0e0e0; }
    .nav-tabs { display: flex; gap: 24px; }
    .nav-tab { padding: 12px 0; font-size: 14px; font-weight: 500; color: #5f6368; cursor: pointer; position: relative; }
    .nav-tab:hover { color: #1a73e8; }
    .nav-tab.active { color: #1a73e8; }
    .nav-tab.active::after { content: ''; position: absolute; bottom: -1px; left: 0; right: 0; height: 3px; background: #1a73e8; border-radius: 3px 3px 0 0; }
    
    .view-controls { display: flex; align-items: center; background: #f1f3f4; border-radius: 4px; padding: 2px; margin-bottom: 8px; }
    .vc-btn { padding: 4px 8px; cursor: pointer; border-radius: 4px; color: #5f6368; display: flex; align-items: center; justify-content: center; transition: background 0.2s; }
    .vc-btn.active { background: #ffffff; color: #202124; box-shadow: 0 1px 2px rgba(0,0,0,0.1); }
    .vc-btn .material-symbols-outlined { font-size: 18px; }

    /* Grid View */
    .grid-container { margin-top: 16px; }
    .grid-section-title { font-size: 13px; font-weight: 600; color: #5f6368; margin-bottom: 16px; }
    .grid-layout { display: flex; gap: 16px; flex-wrap: wrap; }
    .grid-card { width: 220px; border: 1px solid #e0e0e0; border-radius: 8px; padding: 16px; cursor: pointer; transition: box-shadow 0.2s, border-color 0.2s; background: #fff; display: flex; flex-direction: column; gap: 12px; }
    .grid-card:hover { box-shadow: 0 4px 12px rgba(0,0,0,0.08); border-color: #dadce0; }
    .gc-header { display: flex; align-items: center; gap: 8px; }
    .gc-star { color: #9aa0a6; font-size: 18px; cursor: pointer; }
    .gc-star.is-fav { color: #fbbc04; }
    .gc-icon { font-size: 20px; }
    .gc-title { font-size: 14px; font-weight: 500; color: #202124; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; flex: 1; }
    .gc-footer { font-size: 12px; color: #5f6368; padding-left: 26px; }

    /* List Container */
    .list-container { margin-top: 16px; }
    .list-header { display: flex; align-items: center; padding: 12px 16px; border-bottom: 1px solid #e0e0e0; font-size: 13px; font-weight: 600; color: #5f6368; }
    .list-row { display: flex; align-items: center; padding: 12px 16px; border-bottom: 1px solid #e0e0e0; cursor: pointer; transition: background 0.2s; }
    .list-row:hover { background: #f8f9fa; }
    
    .col-name { flex: 2; display: flex; align-items: center; min-width: 0; }
    .list-header .col-name { color: #5f6368; }
    
    .fav-icon { color: #9aa0a6; font-size: 18px; margin-right: 12px; cursor: pointer; }
    .fav-icon.is-fav { color: #fbbc04; }
    
    .doc-icon { font-size: 20px; margin-right: 16px; }
    .doc-title-text { font-size: 14px; font-weight: 500; color: #202124; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    
    .col-owner { flex: 1; display: flex; align-items: center; gap: 8px; color: #5f6368; }
    .owner-avatar { width: 24px; height: 24px; border-radius: 50%; background: #1a73e8; color: #fff; display: flex; align-items: center; justify-content: center; font-size: 10px; font-weight: 600; }
    .owner-name { font-size: 13px; }
    
    .col-date { flex: 1; font-size: 13px; color: #5f6368; }
    
    .col-actions { width: 80px; display: flex; align-items: center; justify-content: flex-end; gap: 4px; opacity: 0; transition: opacity 0.2s; }
    .list-row:hover .col-actions { opacity: 1; }
    .action-btn { background: none; border: none; color: #5f6368; padding: 6px; border-radius: 50%; cursor: pointer; display: flex; align-items: center; justify-content: center; transition: background 0.2s; }
    .action-btn .material-symbols-outlined { font-size: 18px; }
    .action-btn:hover { background: #e8eaed; color: #202124; }
    
    .empty-state { padding: 64px 24px; text-align: center; color: #5f6368; font-size: 15px; }
    .shadow-lg { box-shadow: 0 4px 12px rgba(0,0,0,0.15); }
    
    .toast { position: fixed; bottom: 32px; left: 50%; transform: translateX(-50%) translateY(20px); background: #323232; color: #fff; padding: 12px 24px; border-radius: 4px; font-size: 14px; opacity: 0; transition: all .25s; pointer-events: none; z-index: 1000; }
    .toast.show { opacity: 1; transform: translateX(-50%) translateY(0); }

    /* Modals */
    .modal-backdrop { position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.5); z-index: 2000; display: flex; align-items: center; justify-content: center; }
    .delete-modal { background: #fff; border-radius: 8px; width: 400px; max-width: 90vw; display: flex; flex-direction: column; overflow: hidden; animation: modalIn 0.2s ease-out; }
    .dm-header { padding: 16px 24px; border-bottom: 1px solid #e0e0e0; display: flex; align-items: center; gap: 12px; font-size: 16px; font-weight: 600; color: #202124; }
    .dm-body { padding: 24px; font-size: 14px; color: #3c4043; line-height: 1.5; margin: 0; }
    .dm-body p { margin: 0; }
    .dm-footer { padding: 16px 24px; background: #f8f9fa; border-top: 1px solid #e0e0e0; display: flex; justify-content: flex-end; gap: 12px; }
    .btn-danger { background: #d32f2f; color: #fff; border: none; }
    .btn-danger:hover { background: #b71c1c; box-shadow: 0 1px 3px rgba(0,0,0,0.2); }
    @keyframes modalIn { from { opacity: 0; transform: scale(0.95); } to { opacity: 1; transform: scale(1); } }
    
    /* Upload Overlay (2026 Premium) */
    .upload-overlay { position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(255, 255, 255, 0.85); backdrop-filter: blur(8px); -webkit-backdrop-filter: blur(8px); z-index: 9999; display: flex; align-items: center; justify-content: center; animation: fadeIn 0.3s ease-out; }
    .upload-modal { background: #fff; border-radius: 16px; padding: 40px; width: 420px; text-align: center; border: 1px solid rgba(0,0,0,0.08); box-shadow: 0 20px 40px rgba(0,0,0,0.1); animation: slideUp 0.4s cubic-bezier(0.16, 1, 0.3, 1); }
    .um-icon { width: 64px; height: 64px; border-radius: 50%; background: #e8f0fe; display: flex; align-items: center; justify-content: center; margin: 0 auto 20px auto; }
    .um-title { font-size: 20px; font-weight: 600; color: #202124; margin-bottom: 8px; letter-spacing: -0.3px; }
    .um-subtitle { font-size: 14px; color: #5f6368; line-height: 1.5; margin-bottom: 28px; padding: 0 20px; }
    .um-progress-container { height: 8px; background: #e0e0e0; border-radius: 4px; overflow: hidden; margin-bottom: 16px; position: relative; }
    .um-progress-bar { height: 100%; background: linear-gradient(90deg, #1a73e8, #4285f4); border-radius: 4px; transition: width 0.2s ease-out; position: relative; overflow: hidden; }
    .um-progress-bar::after { content: ''; position: absolute; top: 0; left: 0; right: 0; bottom: 0; background: linear-gradient(90deg, rgba(255,255,255,0) 0%, rgba(255,255,255,0.3) 50%, rgba(255,255,255,0) 100%); animation: shimmer 1.5s infinite; }
    .um-stats { display: flex; justify-content: space-between; font-size: 13px; font-weight: 500; }
    .um-percent { color: #1a73e8; font-weight: 600; font-size: 14px; }
    .um-time { color: #5f6368; }
    
    @keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }
    @keyframes slideUp { from { opacity: 0; transform: translateY(20px) scale(0.98); } to { opacity: 1; transform: translateY(0) scale(1); } }
    @keyframes shimmer { 0% { transform: translateX(-100%); } 100% { transform: translateX(100%); } }
  `]
})
export class DashboardComponent implements OnInit {
  docs: any[] = [];
  toastVisible = false;
  toastMsg = '';

  menus: { [key: string]: boolean } = {
    launcher: false,
    newMenu: false,
    notif: false,
    profile: false,
    apps: false,
    upload: false
  };

  currentApp = 'Show';
  currentTab = 'My Documents';
  viewMode = 'list';
  searchQuery = '';
  searchFocused = false;

  activeWidget: string | null = null;
  notesContent = '';
  scratchpadContent = '';

  filterDate: Date | null = null;
  calMonth = new Date().getMonth();
  calYear = new Date().getFullYear();

  get currentMonthName() { return new Date(this.calYear, this.calMonth).toLocaleString('default', { month: 'long' }); }
  get calendarDays() {
    const days: (Date | null)[] = [];
    const firstDay = new Date(this.calYear, this.calMonth, 1).getDay();
    const daysInMonth = new Date(this.calYear, this.calMonth + 1, 0).getDate();
    for (let i = 0; i < firstDay; i++) days.push(null);
    for (let i = 1; i <= daysInMonth; i++) days.push(new Date(this.calYear, this.calMonth, i));
    return days;
  }
  changeMonth(dir: number) {
    this.calMonth += dir;
    if (this.calMonth < 0) { this.calMonth = 11; this.calYear--; }
    else if (this.calMonth > 11) { this.calMonth = 0; this.calYear++; }
  }
  isSameDate(d1: Date, d2: Date | null) {
    if (!d2) return false;
    return d1.getFullYear() === d2.getFullYear() && d1.getMonth() === d2.getMonth() && d1.getDate() === d2.getDate();
  }
  hasDocsOnDate(d: Date) {
    return this.docs.some(doc => {
      const docDate = new Date(doc.updated_at);
      return docDate.getFullYear() === d.getFullYear() && docDate.getMonth() === d.getMonth() && docDate.getDate() === d.getDate();
    });
  }
  selectDate(d: Date) {
    if (this.isSameDate(d, this.filterDate)) this.filterDate = null;
    else this.filterDate = d;
  }
  toggleWidget(w: string) {
    if (this.activeWidget === w) this.activeWidget = null;
    else { this.activeWidget = w; this.menus['launcher'] = false; this.menus['apps'] = false; }
  }

  get initials() {
    return (this.auth.user?.name ?? 'U').split(' ').map(w => w[0]).join('').toUpperCase().slice(0, 2);
  }

  private _lastDocs: any[] = [];
  private _lastFilterDate: Date | null = null;
  private _lastSearchQuery = '';
  private _lastCurrentApp = '';
  private _lastCurrentTab = '';
  private _lastUserId: string | undefined = undefined;
  private _cachedFilteredDocs: any[] = [];

  get filteredDocs() {
    const currentUserId = this.auth.user?.id;
    if (this.docs === this._lastDocs &&
        this.filterDate === this._lastFilterDate &&
        this.searchQuery === this._lastSearchQuery &&
        this.currentApp === this._lastCurrentApp &&
        this.currentTab === this._lastCurrentTab &&
        currentUserId === this._lastUserId) {
      return this._cachedFilteredDocs;
    }

    let list = this.docs;

    if (this.filterDate) {
      list = list.filter(d => {
        const docDate = new Date(d.updated_at);
        return this.isSameDate(docDate, this.filterDate);
      });
    }

    // Filter by Search Query First
    if (this.searchQuery) {
      const q = this.searchQuery.toLowerCase();
      list = list.filter(d => d.title.toLowerCase().includes(q));
    }

    // Filter by App
    if (this.currentApp === 'Writer') list = list.filter(d => d.doc_type === 'doc');
    if (this.currentApp === 'Sheet') list = list.filter(d => d.doc_type === 'sheet');
    if (this.currentApp === 'Show') list = list.filter(d => d.doc_type === 'slide');

    // Filter by Tab
    if (this.currentTab === 'Trash') {
      list = list.filter(d => d.is_trashed === 1 || d.is_trashed === true || d.is_trashed == '1');
    } else {
      list = list.filter(d => !d.is_trashed || d.is_trashed === 0 || d.is_trashed == '0');
      if (this.currentTab === 'Favorites') {
        list = list.filter(d => d._favorite);
      } else if (this.currentTab === 'Shared with me') {
        const myId = this.auth.user?.id;
        list = list.filter(d => myId != null && String(d.owner_id) !== String(myId));
      } else if (this.currentTab === 'My Documents') {
        const myId = this.auth.user?.id;
        list = list.filter(d => myId != null && String(d.owner_id) === String(myId));
      } else if (this.currentTab === 'Recents') {
        const twoDaysAgo = new Date();
        twoDaysAgo.setDate(twoDaysAgo.getDate() - 2);
        list = list.filter(d => new Date(d.updated_at) >= twoDaysAgo).slice(0, 10);
      }
    }
    
    this._lastDocs = this.docs;
    this._lastFilterDate = this.filterDate;
    this._lastSearchQuery = this.searchQuery;
    this._lastCurrentApp = this.currentApp;
    this._lastCurrentTab = this.currentTab;
    this._lastUserId = currentUserId;
    this._cachedFilteredDocs = list;
    
    return list;
  }

  private _lastDocsForResults: any[] = [];
  private _lastSearchQueryForResults = '';
  private _cachedSearchResults: any[] = [];

  get searchResults() {
    if (this.docs === this._lastDocsForResults && this.searchQuery === this._lastSearchQueryForResults) {
      return this._cachedSearchResults;
    }
    
    let res: any[] = [];
    if (!this.searchQuery) res = this.docs.filter(d => !d.is_trashed).slice(0, 5);
    else {
      const q = this.searchQuery.toLowerCase();
      res = this.docs.filter(d => d.title.toLowerCase().includes(q) && !d.is_trashed).slice(0, 5);
    }
    
    this._lastDocsForResults = this.docs;
    this._lastSearchQueryForResults = this.searchQuery;
    this._cachedSearchResults = res;
    return res;
  }

  isOwner(doc: any): boolean {
    if (!this.auth.user || !this.auth.user.id) return false;
    return String(doc.owner_id) === String(this.auth.user.id);
  }

  constructor(public auth: AuthService, private api: ApiService, private router: Router) { }

  ngOnInit() {
    this.detectSubdomain();
    this.auth.ready$.pipe(
      filter(ready => ready),
      take(1)
    ).subscribe(() => this.load());
  }

  detectSubdomain() {
    const host = window.location.hostname;
    if (host.includes('docs') || host.includes('writer')) {
      this.currentApp = 'Writer';
    } else if (host.includes('sheet')) {
      this.currentApp = 'Sheet';
    } else if (host.includes('show') || host.includes('slide')) {
      this.currentApp = 'Show';
    } else {
      this.currentApp = 'Show';
    }
    // ── Sync browser tab title + favicon to the current app ──────────────────
    const faviconMap: Record<string, {title: string; svg: string}> = {
      Sheet: {
        title: 'Vsnap Sheet',
        svg: `data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'><rect width='32' height='32' rx='5' fill='%2326A96C'/><rect x='5' y='8' width='22' height='2.5' rx='1.2' fill='white'/><rect x='5' y='13.5' width='22' height='2.5' rx='1.2' fill='white'/><rect x='5' y='19' width='14' height='2.5' rx='1.2' fill='white'/></svg>`,
      },
      Writer: {
        title: 'Vsnap Writer',
        svg: `data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'><rect width='32' height='32' rx='5' fill='%234285F4'/><rect x='7' y='7' width='18' height='2.5' rx='1.2' fill='white'/><rect x='7' y='12' width='18' height='2.5' rx='1.2' fill='white'/><rect x='7' y='17' width='18' height='2.5' rx='1.2' fill='white'/><rect x='7' y='22' width='11' height='2.5' rx='1.2' fill='white'/></svg>`,
      },
      Show: {
        title: 'Vsnap Show',
        svg: `data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'><rect width='32' height='32' rx='5' fill='%23F4A623'/><rect x='4' y='6' width='24' height='16' rx='2' fill='white' opacity='0.9'/><rect x='6' y='8' width='20' height='12' rx='1.5' fill='%23F4A623'/><rect x='14' y='22' width='4' height='4' rx='1' fill='white' opacity='0.9'/><rect x='10' y='26' width='12' height='2' rx='1' fill='white' opacity='0.7'/></svg>`,
      },
      WorkDrive: {
        title: 'Vsnap WorkDrive',
        svg: `data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'><rect width='32' height='32' rx='5' fill='%236366F1'/><rect x='4' y='10' width='24' height='14' rx='2' fill='white' opacity='0.9'/><rect x='4' y='10' width='10' height='4' rx='1.5' fill='%236366F1'/><rect x='4' y='8' width='14' height='4' rx='2' fill='white'/></svg>`,
      },
    };
    const app = faviconMap[this.currentApp] || faviconMap['WorkDrive'];
    document.title = app.title;
    let link: HTMLLinkElement = document.querySelector("link[rel~='icon']") as HTMLLinkElement;
    if (!link) { link = document.createElement('link'); link.rel = 'icon'; document.head.appendChild(link); }
    link.href = app.svg;
    // ─────────────────────────────────────────────────────────────────
  }

  load() {
    this.api.listDocuments().subscribe(d => {
      let favs = JSON.parse(localStorage.getItem('vsnap_favs') || '{}');
      this.docs = d.map(doc => ({ ...doc, _favorite: !!favs[doc.id] }));
    });
  }

  setApp(app: string, event: Event) {
    event.stopPropagation();
    this.closeAllMenus();
    
    if (app === 'WorkDrive') {
      this.currentApp = app;
      return;
    }
    
    const env = window.location.hostname.startsWith('test') ? 'test' : '';
    let domain = '';
    
    if (app === 'Writer') domain = `${env}docs`;
    else if (app === 'Sheet') domain = `${env}sheets`;
    else if (app === 'Show') domain = `${env}show`;

    if (domain) {
      window.location.href = `https://${domain}.vsnaptechnology.com/`;
    }
  }

  toggleMenu(menuName: string, event: Event) {
    event.stopPropagation();
    const current = this.menus[menuName];
    this.closeAllMenus();
    this.menus[menuName] = !current;
  }

  closeAllMenus() {
    Object.keys(this.menus).forEach(k => this.menus[k] = false);
    this.searchFocused = false;
  }

  onSearchBlur() {
    setTimeout(() => this.searchFocused = false, 200);
  }

  toggleFavorite(doc: any, event: Event) {
    event.stopPropagation();
    doc._favorite = !doc._favorite;

    let favs = JSON.parse(localStorage.getItem('vsnap_favs') || '{}');
    favs[doc.id] = doc._favorite;
    localStorage.setItem('vsnap_favs', JSON.stringify(favs));

    if (doc._favorite) {
      this.showToast('Added to Favorites');
    } else {
      this.showToast('Removed from Favorites');
    }
  }

  isUploading = false;
  uploadProgress = 0;
  uploadTimeLeft = '';
  private uploadStartTime = 0;

  onUpload(event: any) {
    const file = event.target.files[0];
    if (file) {
      this.isUploading = true;
      this.uploadProgress = 0;
      this.uploadTimeLeft = 'Calculating...';
      this.uploadStartTime = Date.now();
      
      this.api.importFile(file).subscribe({
        next: (event: any) => {
          if (event.type === HttpEventType.UploadProgress) {
            if (event.total) {
              this.uploadProgress = Math.round(100 * event.loaded / event.total);
              const timeElapsed = Date.now() - this.uploadStartTime;
              const uploadSpeed = event.loaded / (timeElapsed / 1000); // bytes per sec
              const bytesLeft = event.total - event.loaded;
              const secondsLeft = Math.round(bytesLeft / uploadSpeed);
              
              if (secondsLeft > 60) {
                this.uploadTimeLeft = `~${Math.ceil(secondsLeft/60)} mins left`;
              } else if (secondsLeft > 0) {
                this.uploadTimeLeft = `${secondsLeft} secs left`;
              } else {
                this.uploadTimeLeft = 'Processing file...';
              }
            }
          } else if (event.type === HttpEventType.Response) {
            this.showToast(`${file.name} uploaded successfully.`);
            this.load();
            this.isUploading = false;
          }
        },
        error: (err: any) => {
          this.isUploading = false;
          this.showToast(`Error uploading ${file.name}.`);
          console.error('Import failed:', err);
        }
      });
    }
  }

  create(type: string) {
    this.closeAllMenus();
    this.api.createDocument('Untitled', type).subscribe((doc: any) => {
      this.router.navigate([`/${doc.doc_type}/${doc.id}`]);
    });
  }

  open(doc: any) {
    this.closeAllMenus();
    this.router.navigate([`/${doc.doc_type}/${doc.id}`]);
  }

  isDeleting: { [key: string]: boolean } = {};
  deleteConfirmDoc: any = null;
  deleteConfirmPermanent = false;

  delete(doc: any) {
    this.deleteConfirmPermanent = doc.is_trashed === 1 || doc.is_trashed === true || doc.is_trashed == '1';
    this.deleteConfirmDoc = doc;
  }

  confirmDelete() {
    if (!this.deleteConfirmDoc) return;
    const doc = this.deleteConfirmDoc;
    const isPermanent = this.deleteConfirmPermanent;

    if (this.isDeleting[doc.id]) return;
    this.isDeleting[doc.id] = true;
    this.api.deleteDocument(doc.id).subscribe({
      next: () => {
        this.showToast(isPermanent ? 'Permanently deleted.' : 'Moved to trash.');
        this.load();
        delete this.isDeleting[doc.id];
        this.deleteConfirmDoc = null;
      },
      error: () => {
        delete this.isDeleting[doc.id];
        this.deleteConfirmDoc = null;
      }
    });
  }

  cancelDelete() {
    this.deleteConfirmDoc = null;
  }

  restore(doc: any, event: Event) {
    event.stopPropagation();
    this.api.restoreDocument(doc.id).subscribe(() => {
      this.showToast('Document restored.');
      this.load();
    });
  }

  trackById(index: number, doc: any) {
    return doc.id;
  }

  copyLink(doc: any) {
    const url = `${window.location.origin}/${doc.doc_type}/${doc.id}`;
    navigator.clipboard.writeText(url).then(() => this.showToast('Link copied! Share it with anyone.'));
  }

  formatDate(d: string) {
    if (!d) return '';
    return new Date(d).toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: 'numeric' });
  }

  showToast(msg: string) {
    this.toastMsg = msg; this.toastVisible = true;
    setTimeout(() => this.toastVisible = false, 2500);
  }
}