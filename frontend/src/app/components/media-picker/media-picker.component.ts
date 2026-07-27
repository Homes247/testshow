import { Component, EventEmitter, Input, Output, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { PickerModule } from '@ctrl/ngx-emoji-mart';
import { GiphyService } from '../../services/giphy.service';

@Component({
  selector: 'app-media-picker',
  standalone: true,
  imports: [CommonModule, FormsModule, PickerModule],
  template: `
    <div class="media-picker-container" [class.dark]="darkMode" (click)="$event.stopPropagation()">
      <!-- Tabs -->
      <div class="mp-tabs">
        <button class="mp-tab" [class.active]="activeTab === 'emoji'" (click)="activeTab = 'emoji'">
          <span class="material-symbols-outlined">mood</span> Emojis
        </button>
        <button class="mp-tab" [class.active]="activeTab === 'gif'" (click)="loadTrendingGifs()">
          <span class="material-symbols-outlined">gif_box</span> GIFs
        </button>
        <button class="mp-tab" [class.active]="activeTab === 'sticker'" (click)="loadTrendingStickers()">
          <span class="material-symbols-outlined">sticky_note_2</span> Stickers
        </button>
      </div>
      
      <!-- Content -->
      <div class="mp-content">
        <!-- Emojis -->
        <div class="mp-pane" *ngIf="activeTab === 'emoji'">
          <emoji-mart (emojiSelect)="onEmojiSelect($event)" [emojisToShowFilter]="filterEmojis" title="Pick your emoji…" emoji="smile" [style]="{ width: '100%', height: pickerHeight, border: 'none', padding: '0' }" set="apple" [darkMode]="darkMode"></emoji-mart>
        </div>

        <!-- GIFs -->
        <div class="mp-pane mp-media-pane" *ngIf="activeTab === 'gif'">
          <div class="mp-search">
            <span class="material-symbols-outlined">search</span>
            <input type="text" [(ngModel)]="gifSearchQuery" (ngModelChange)="onGifSearch()" placeholder="Search GIFs..." />
          </div>
          <div class="mp-grid" [style.height]="gridHeight">
            <div class="mp-loading" *ngIf="loading">Loading...</div>
            <img class="mp-grid-item" *ngFor="let g of gifs" [src]="g.images.fixed_height_small.url" (click)="onMediaSelect(g.images.original.url, 'gif')" loading="lazy" />
          </div>
        </div>

        <!-- Stickers -->
        <div class="mp-pane mp-media-pane" *ngIf="activeTab === 'sticker'">
          <div class="mp-search">
            <span class="material-symbols-outlined">search</span>
            <input type="text" [(ngModel)]="stickerSearchQuery" (ngModelChange)="onStickerSearch()" placeholder="Search Stickers..." />
          </div>
          <div class="mp-grid" [style.height]="gridHeight">
            <div class="mp-loading" *ngIf="loading">Loading...</div>
            <img class="mp-grid-item" *ngFor="let s of stickers" [src]="s.images.fixed_height_small.url" (click)="onMediaSelect(s.images.original.url, 'sticker')" loading="lazy" />
          </div>
        </div>
      </div>
    </div>
  `,
  styles: [`
    .media-picker-container { display:flex; flex-direction:column; width:100%; height:100%; background: white; border-top: 1px solid #e2e8f0; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.15); overflow: hidden; }
    .media-picker-container.dark { background: #1e293b; border-top-color: #334155; }
    .mp-tabs { display:flex; border-bottom: 1px solid #e2e8f0; }
    .media-picker-container.dark .mp-tabs { border-bottom-color: #334155; }
    .mp-tab { flex:1; display:flex; align-items:center; justify-content:center; gap:6px; padding:10px; background:none; border:none; cursor:pointer; font-family:inherit; font-size:12px; font-weight:600; color:#64748b; border-bottom: 2px solid transparent; transition: all 0.2s; }
    .media-picker-container.dark .mp-tab { color:#94a3b8; }
    .mp-tab:hover { color:#0f172a; background:#f8fafc; }
    .media-picker-container.dark .mp-tab:hover { color:#e2e8f0; background:#0f172a; }
    .mp-tab.active { color:#0E7ACA; border-bottom-color:#0E7ACA; }
    .mp-tab .material-symbols-outlined { font-size:18px; }
    .mp-content { flex:1; overflow:hidden; display:flex; flex-direction:column; }
    .mp-pane { flex:1; display:flex; flex-direction:column; overflow:hidden; }
    .mp-media-pane { padding: 8px; }
    .mp-search { display:flex; align-items:center; gap:8px; padding:6px 12px; background:#f1f5f9; border-radius:8px; margin-bottom:8px; flex-shrink: 0; }
    .media-picker-container.dark .mp-search { background:#0f172a; }
    .mp-search input { flex:1; background:transparent; border:none; outline:none; font-family:inherit; font-size:13px; color:#0f172a; }
    .media-picker-container.dark .mp-search input { color:#e2e8f0; }
    .mp-search .material-symbols-outlined { font-size:16px; color:#94a3b8; }
    .mp-grid { display:grid; grid-template-columns: repeat(auto-fill, minmax(80px, 1fr)); grid-auto-rows: 80px; gap:6px; overflow-y:auto; padding-right:4px; }
    .mp-grid-item { width:100%; height:100%; object-fit:cover; border-radius:6px; cursor:pointer; transition:transform 0.15s; }
    .mp-grid-item:hover { transform:scale(1.05); z-index:2; position:relative; box-shadow:0 4px 12px rgba(0,0,0,0.2); }
    .mp-loading { grid-column: 1 / -1; text-align:center; padding: 20px; color:#94a3b8; font-size:13px; font-weight:500; }
    
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
  `]
})
export class MediaPickerComponent implements OnInit {
  @Input() darkMode = false;
  @Input() pickerHeight = '250px';
  @Input() gridHeight = '200px';

  @Output() emojiSelect = new EventEmitter<string>();
  @Output() mediaSelect = new EventEmitter<{url: string, type: 'gif' | 'sticker'}>();

  activeTab: 'emoji' | 'gif' | 'sticker' = 'emoji';

  gifSearchQuery = '';
  stickerSearchQuery = '';
  gifs: any[] = [];
  stickers: any[] = [];
  loading = false;
  searchTimeout: any;

  constructor(private giphy: GiphyService) {}

  ngOnInit() {
    // Clear the Frequently Used cache once to remove any previously clicked 
    // unsupported emojis (like ZWJ sequences) that were stored before the filter was added.
    if (!localStorage.getItem('emoji_cache_cleared_v1')) {
      localStorage.removeItem('emoji-mart.frequently');
      localStorage.removeItem('emoji-mart.last');
      localStorage.setItem('emoji_cache_cleared_v1', 'true');
    }
  }

  /**
   * Filter out emojis that Windows 10 cannot render (Unicode 14.0+ and some ZWJ sequences).
   * These show as broken/garbled characters in text inputs on Windows 10.
   */
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

  onEmojiSelect(event: any) {
    // Derive emoji from unified hex code for reliable cross-platform rendering.
    let emojiStr = '';
    if (event.emoji.unified) {
      try {
        const parts = event.emoji.unified.split('-').map((u: string) => parseInt(u, 16));
        emojiStr = String.fromCodePoint(...parts);
      } catch (e) {
        emojiStr = event.emoji.native || '';
      }
    } else {
      emojiStr = event.emoji.native || '';
    }
    this.emojiSelect.emit(emojiStr);
  }

  onMediaSelect(url: string, type: 'gif' | 'sticker') {
    this.mediaSelect.emit({url, type});
  }

  loadTrendingGifs() {
    this.activeTab = 'gif';
    if (this.gifs.length === 0) {
      this.loading = true;
      this.giphy.getTrendingGifs(30).subscribe({
        next: (res) => { this.gifs = res.data; this.loading = false; },
        error: () => { this.loading = false; }
      });
    }
  }

  loadTrendingStickers() {
    this.activeTab = 'sticker';
    if (this.stickers.length === 0) {
      this.loading = true;
      this.giphy.getTrendingStickers(30).subscribe({
        next: (res) => { this.stickers = res.data; this.loading = false; },
        error: () => { this.loading = false; }
      });
    }
  }

  onGifSearch() {
    clearTimeout(this.searchTimeout);
    this.searchTimeout = setTimeout(() => {
      if (!this.gifSearchQuery.trim()) {
        this.gifs = [];
        this.loadTrendingGifs();
        return;
      }
      this.loading = true;
      this.giphy.searchGifs(this.gifSearchQuery.trim(), 30).subscribe({
        next: (res) => { this.gifs = res.data; this.loading = false; },
        error: () => { this.loading = false; }
      });
    }, 500);
  }

  onStickerSearch() {
    clearTimeout(this.searchTimeout);
    this.searchTimeout = setTimeout(() => {
      if (!this.stickerSearchQuery.trim()) {
        this.stickers = [];
        this.loadTrendingStickers();
        return;
      }
      this.loading = true;
      this.giphy.searchStickers(this.stickerSearchQuery.trim(), 30).subscribe({
        next: (res) => { this.stickers = res.data; this.loading = false; },
        error: () => { this.loading = false; }
      });
    }, 500);
  }
}
