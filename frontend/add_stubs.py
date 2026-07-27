import re

filepath = 'src/app/pages/sheet-editor/sheet-editor.component.ts'
content = open(filepath, encoding='utf-8').read()

# All missing members identified from build output
missing_members = """
  // ---- Auto-generated stubs for missing template bindings ----
  activeModal: string | null = null;
  activeShapeIdx: number | null = null;
  activeShapeMenuIdx: number | null = null;
  addPersonalDictWord() {}
  addPicklistOption() {}
  addShareEmail() {}
  addSheet() {}
  addTask() {}
  advSort: any = {};
  allAdvFilterSelected = true;
  applyAdvFilter() {}
  applyCustomFormat() {}
  applyFont(f: string) {}
  applyGoalSeek() {}
  applyPivot() {}
  applyPresetPicklist(p: any) {}
  applySpellCheckFix(i: number) {}
  applyTextToColumns() {}
  applyTranslation() {}
  autoResizeEditor() {}
  back() { this.router.navigate(['/']); }
  cancelRangePicker() {}
  cellHasContent(r: number, c: number): boolean { return !!(this.cells[r] && this.cells[r][c]); }
  clearAdvFilter() {}
  clearAll() { this.clearRangeData(); }
  clearAllFilters() { this.filterActive = false; this.hiddenRows.clear(); if (this.cdr) this.cdr.markForCheck(); }
  clearAllFormats() { this.forEachSelectedCell((r, c) => { delete this.formats[`${r},${c}`]; }); this.onCellChange(); }
  clearCheckboxes() {}
  clearConditionalFormats() {}
  clearDataValidations() {}
  clearGroups() {}
  clearHyperlinks() {}
  clearInlineSearch() { this.inlineSearchQuery = ''; this.inlineSearchMatches = []; this.inlineSearchIdx = 0; }
  clearNotes() {}
  clearRichTextFormats() {}
  closeAdvFilter() { this.advFilterVisible = false; }
  closeConfirm() { this.confirmModalOpen = false; }
  closePrompt() { this.promptModalOpen = false; }
  closeSidePanel() {}
  colGroupMarginHeight = 0;
  commitEdit() { this.isEditingCell = false; if (this.cdr) this.cdr.markForCheck(); }
  commitFormula() { this.isEditingCell = false; this.cells[this.selectedRow][this.selectedCol] = this.formulaBarValue; this.onCellChange(); this.save(); }
  confirmMoveSheet() { this.moveSheetModalOpen = false; }
  confirmRangePicker() {}
  copyLink() { navigator.clipboard.writeText(window.location.href).catch(() => {}); }
  copyPublishLink() {}
  copySheet() {}
  createForm() {}
  createPivotTable() {}
  createSparkline() {}
  customFormatString = '';
  customFunctionsScript = '';
  customInsertCol() {}
  dataFromPicture() {}
  decreaseDecimals() {}
  decrementFontSize() { this.currentSizeNum = Math.max(6, this.currentSizeNum - 1); this.currentSize = this.currentSizeNum + 'px'; }
  defineName() {}
  deleteCol() {}
  deleteRow() {}
  deleteShape() {}
  deleteSheet(idx: number) { if (this.sheets.length > 1) { this.sheets.splice(idx, 1); if (this.currentSheetIdx >= this.sheets.length) this.currentSheetIdx = this.sheets.length - 1; } }
  deleteShiftLeft() {}
  deleteShiftUp() {}
  deleteSparklineConfig() {}
  diagramCategory = 'basic';
  dummyList: any[] = [];
  duplicateSheet() {}
  editHistoryCell = '';
  editHistoryData: any[] = [];
  editShapeLabel() {}
  editSparkline() {}
  emailNotifEmail = '';
  emailNotifModalOpen = false;
  emailNotifOnComment = true;
  emailNotifOnEdit = true;
  executeGoto() {}
  exportFile(fmt: string) {}
  feedbackModalOpen = false;
  feedbackRating = 0;
  feedbackText = '';
  fillLeft() {}
  fillUp() {}
  filterByCellColor() {}
  filterByCellValue() {}
  filterByTextColor() {}
  findNext() {}
  formData: any[] = [];
  formHeaders: string[] = [];
  freezeCols(n: number) {}
  freezeRows(n: number) {}
  freezeSelection() {}
  generateChart() {}
  getCellStyle(r: number, c: number): any { const fmt = this.formats[`${r},${c}`] as any; if (!fmt) return {}; const s: any = {}; if (fmt.bold) s['font-weight'] = 'bold'; if (fmt.italic) s['font-style'] = 'italic'; if (fmt.underline) s['text-decoration'] = 'underline'; if (fmt.color) s['color'] = fmt.color; if (fmt.bg) s['background-color'] = fmt.bg; return s; }
  getColGroupsFor(c: number): any[] { return (this.sheets[this.currentSheetIdx]?.colGroups || []).map((g: any, i: number) => ({ ...g, index: i })).filter((g: any) => c >= g.start && c <= g.end); }
  getColSpan(r: number, c: number): number { return 1; }
  getColWidth(c: number): number { return (this.sheets[this.currentSheetIdx]?.colWidths || {})[c] || 100; }
  getContentStyle(r: number, c: number): any { return {}; }
  getDropdownColor(r: number, c: number): string { return ''; }
  getFormat(r: number, c: number, prop: string): any { return (this.formats[`${r},${c}`] as any)?.[prop]; }
  getFormatName(): string { return (this.formats[`${this.selectedRow},${this.selectedCol}`] as any)?.numFormat || 'general'; }
  getFormatWrap(): boolean { return !!(this.formats[`${this.selectedRow},${this.selectedCol}`] as any)?.wrap; }
  getFrozenColOffset(): number { return 0; }
  getFrozenRowOffset(): number { return 0; }
  getRangePickerValue(): string { return ''; }
  getRowGroupsFor(r: number): any[] { return (this.sheets[this.currentSheetIdx]?.rowGroups || []).map((g: any, i: number) => ({ ...g, index: i })).filter((g: any) => r >= g.start && r <= g.end); }
  getRowSpan(r: number, c: number): number { return 1; }
  getSparklineSvgSafe(r: number, c: number): any { return null; }
  getStatsFilledCells(): number { let n = 0; for (let r = 0; r < this.ROWS; r++) for (let c = 0; c < this.COLS; c++) if (this.cells[r]?.[c]) n++; return n; }
  getStatsFormulaCells(): number { let n = 0; for (let r = 0; r < this.ROWS; r++) for (let c = 0; c < this.COLS; c++) if (this.cells[r]?.[c]?.startsWith('=')) n++; return n; }
  getStatsLockedSheets(): number { return this.sheets.filter((s: any) => s.locked).length; }
  getStatsNumericCells(): number { let n = 0; for (let r = 0; r < this.ROWS; r++) for (let c = 0; c < this.COLS; c++) { const v = this.cells[r]?.[c]; if (v && !isNaN(Number(v)) && v.trim() !== '') n++; } return n; }
  getTasksDone(): number { return 0; }
  getVisibleSheetCount(): number { return this.sheets.filter((s: any) => !s.hidden).length; }
  goalSeekByCell = '';
  goalSeekModalOpen = false;
  goalSeekTargetCell = '';
  goalSeekTargetValue = 0;
  groupCol() {}
  groupMarginWidth = 0;
  groupRow() {}
  handleModalAction(action: string) {}
  hasCellDropdown(r: number, c: number): boolean { return (this.validations[`${r},${c}`] as any)?.type === 'list'; }
  hasColGroups = false;
  hasRowGroups = false;
  hiddenSheetsList: any[] = [];
  hideCols() {}
  hideRows() {}
  hideSheet(idx: number) { this.sheets[idx].hidden = true; }
  increaseDecimals() {}
  incrementFontSize() { this.currentSizeNum = Math.min(96, this.currentSizeNum + 1); this.currentSize = this.currentSizeNum + 'px'; }
  inlineFindNext() {}
  inlineFindPrev() {}
  inlineSearchIdx = 0;
  inlineSearchMatches: Array<{r: number, c: number}> = [];
  inlineSearchQuery = '';
  insertButton() {}
  insertCheckbox() {}
  insertColLeft() {}
  insertColRight() {}
  insertComment() {}
  insertEmoji(emoji: string) { this.cells[this.selectedRow][this.selectedCol] = (this.cells[this.selectedRow][this.selectedCol] || '') + emoji; this.onCellChange(); }
  insertLink() {}
  insertNote() {}
  insertRowAbove() {}
  insertRowBelow() {}
  insertShape(type: string) {}
  isCellActiveInlineSearch(r: number, c: number): boolean { return this.inlineSearchMatches.length > 0 && this.inlineSearchMatches[this.inlineSearchIdx]?.r === r && this.inlineSearchMatches[this.inlineSearchIdx]?.c === c; }
  isCellInInlineSearch(r: number, c: number): boolean { return this.inlineSearchMatches.some(m => m.r === r && m.c === c); }
  isColActiveAxis(c: number): boolean { return false; }
  isColHeaderSelected(c: number): boolean { return this.selectedColHeader === c; }
  isColumnFiltered(c: number): boolean { return this.activeFilterCols.has(c); }
  isFilterHeaderCell(r: number, c: number): boolean { return this.filterActive && r === 0; }
  isImageCell(r: number, c: number): boolean { const v = this.cells[r]?.[c] || ''; return v.startsWith('__img__'); }
  isMergedSlave(r: number, c: number): boolean { return false; }
  isRowActiveAxis(r: number): boolean { return false; }
  isRowHeaderSelected(r: number): boolean { return this.selectedRowHeader === r; }
  isSparklineCell(r: number, c: number): boolean { return !!((this.sheets[this.currentSheetIdx]?.sparklines || {})[`${r},${c}`]); }
  isStarred = false;
  lastSavedTime: Date | null = null;
  linkSpreadsheet() {}
  lockCurrentSheet() {}
  lockSelectedRange() {}
  macroScript = '';
  makePublic() {}
  menuSearch = '';
  mergeCells() {}
  modalInput = '';
  moveSheetDestination = 0;
  moveSheetModalOpen = false;
  moveSheetTargetIdx = -1;
  myDocs: any[] = [];
  newDoc() {}
  onCellClickWithPicker(r: number, c: number) {}
  onCellRightClick(e: MouseEvent, r: number, c: number) { e.preventDefault(); this.ctxX = e.clientX; this.ctxY = e.clientY; this.ctxVisible = true; this.selectedRow = r; this.selectedCol = c; }
  onEditorKeydown(e: KeyboardEvent) {}
  onEmojiSelect(emoji: any) { const ch = typeof emoji === 'string' ? emoji : (emoji?.emoji || emoji?.native || ''); this.insertEmoji(ch); this.showEmojiPicker = false; }
  onFileSelected(e: Event) {}
  onFontSizeInputChange(e: Event) { const v = parseInt((e.target as HTMLInputElement).value, 10); if (!isNaN(v) && v > 0) { this.currentSizeNum = v; this.currentSize = v + 'px'; } }
  onHeaderRightClick(e: MouseEvent, type: string, idx: number) { e.preventDefault(); this.ctxX = e.clientX; this.ctxY = e.clientY; this.ctxVisible = true; }
  onImageFileSelected(e: Event) {}
  onInlineSearch() { this.inlineSearchMatches = []; if (!this.inlineSearchQuery) return; for (let r = 0; r < this.ROWS; r++) for (let c = 0; c < this.COLS; c++) if ((this.cells[r]?.[c] || '').toLowerCase().includes(this.inlineSearchQuery.toLowerCase())) this.inlineSearchMatches.push({r, c}); this.inlineSearchIdx = 0; }
  onRangePickerInput(e: Event) {}
  onShareSearch() {}
  openApp(app: string) {}
  openAuditTrail() {}
  openCellEditHistory(r: number, c: number) {}
  openCustomFormatModal() { this.showCustomFormatModal = true; }
  openCustomFunctions() {}
  openDataValidationModal() {}
  openDeveloperApi() {}
  openEditHistory() {}
  openEmailNotifications() { this.emailNotifModalOpen = true; }
  openFeatureModal(f: string) {}
  openFeedback() { this.feedbackModalOpen = true; }
  openFilterMenu(e: MouseEvent, c: number) {}
  openGoalSeek() { this.goalSeekModalOpen = true; }
  openMacroEditor() {}
  openManageRulesModal() { this.manageRulesModalOpen = true; }
  openMergeTemplate() {}
  openMoreFormatsModal() { this.showMoreFormatsModal = true; }
  openMoveSheetModal(idx: number) { this.moveSheetTargetIdx = idx; this.moveSheetModalOpen = true; }
  openPivotModal() { this.pivotModalOpen = true; }
  openPreferences() { this.preferencesModalOpen = true; }
  openSheetMenu(e: MouseEvent, idx: number) { e.stopPropagation(); this.activeSheetMenuIdx = idx; this.sheetMenuX = (e.target as HTMLElement).getBoundingClientRect().left; this.sheetMenuY = (e.target as HTMLElement).getBoundingClientRect().bottom; }
  openSolver() {}
  openTextToColumnsModal() { this.textToColsModalOpen = true; }
  openUserGuide() {}
  openValidationModal() { this.validationModalOpen = true; }
  openWhatsNew() {}
  pasteCell() {}
  pasteExceptBorders() {}
  pasteExceptNotes() {}
  pasteFormats() {}
  pasteFormulas() {}
  pasteFormulasAndNumberFormats() {}
  pasteLinkToSource() {}
  pasteNotes() {}
  pasteSheet() {}
  pasteTranspose() {}
  pasteValidation() {}
  pasteValues() {}
  pasteValuesAndNumberFormats() {}
  patternFill: any = null;
  performShare() {}
  personalDictModalOpen = false;
  personalDictNewWord = '';
  personalDictWords: string[] = [];
  personalDictionary: string[] = [];
  pivotConfig: any = {};
  pivotHeaders: string[] = [];
  prefDateFormat = 'MM/DD/YYYY';
  prefLocale = 'en-US';
  prefThousands = true;
  preferencesModalOpen = false;
  previewImageUrl: string | null = null;
  printSheet() { window.print(); }
  publishRange = '';
  publishSheet() {}
  recalculate() { this.updateDisplayCache(); }
  removeDuplicates() {}
  removePersonalDictWord(w: string) {}
  removeShareEmail(e: string) {}
  removeTask(i: number) {}
  removeValidation() { delete this.validations[`${this.selectedRow},${this.selectedCol}`]; this.onCellChange(); }
  renameSheet(idx: number, name: string) { if (this.sheets[idx]) this.sheets[idx].name = name; }
  replaceAll() {}
  replaceOne() {}
  runMacro() {}
  runTranslate() {}
  saveCustomFunctions() {}
  saveDataValidation() {}
  saveEmailNotifications() { this.emailNotifModalOpen = false; }
  savePreferences() { this.preferencesModalOpen = false; }
  saveSparkline() {}
  saveStatus = '';
  saveValidation() {}
  selectEntireCol(c: number) { this.rangeStart = { r: 0, c }; this.rangeEnd = { r: this.ROWS - 1, c }; this.selectedColHeader = c; }
  selectEntireRow(r: number) { this.rangeStart = { r, c: 0 }; this.rangeEnd = { r, c: this.COLS - 1 }; this.selectedRowHeader = r; }
  selectShareUser(u: any) {}
  setBorders(b: any) {}
  setGridDirection(d: string) {}
  setGridSpacing(n: number) {}
  setGridlineColor(c: string) { if (this.sheets[this.currentSheetIdx]) this.sheets[this.currentSheetIdx].gridlineColor = c; }
  setNumFormat(fmt: string) { this.forEachSelectedCell((r, c) => { if (!this.formats[`${r},${c}`]) this.formats[`${r},${c}`] = {} as any; (this.formats[`${r},${c}`] as any).numFormat = fmt; }); this.onCellChange(); }
  setTabColor(idx: number, color: string) { if (this.sheets[idx]) this.sheets[idx].tabColor = color; }
  setZoom(z: number) { this.zoomLevel = z; }
  shapeCategory = 'basic';
  shapeTab = 'shapes';
  shiftCellsDown() {}
  shiftCellsRight() {}
  showAllComments() {}
  showCustomFormatModal = false;
  showEditHistoryPanel = false;
  showKeyboardShortcuts() {}
  showMoreFormatsModal = false;
  showWordCount() {}
  simulateMerge(r: number, c: number): boolean { return false; }
  sortColAZ() {}
  sortColZA() {}
  spellCheck() {}
  spellCheckErrors: any[] = [];
  spellCheckLoading = false;
  spellCheckModalOpen = false;
  startColResize(e: MouseEvent, c: number) {}
  startRangePicker(field: string) {}
  startRowResize(e: MouseEvent, r: number) {}
  startShapeDrag(e: MouseEvent, idx: number) {}
  statsModalOpen = false;
  submitFeedback() { this.feedbackModalOpen = false; }
  submitForm() {}
  submitPrompt() { this.promptModalOpen = false; }
  switchSheet(idx: number) { this.currentSheetIdx = idx; this.cells = this.sheets[idx].cells; this.formats = this.sheets[idx].formats; this.validations = this.sheets[idx].validations || {}; this.hiddenRows = new Set(this.sheets[idx].hiddenRows || []); this.activeFilterCols = new Set(this.sheets[idx].activeFilterCols || []); this.filterActive = !!this.sheets[idx].filterActive; this.updateDisplayCache(); }
  t2cCustomDelim = '';
  t2cDelimiter = 'comma';
  textCategory = 'basic';
  textToColsModalOpen = false;
  toggleAllAdvFilter(v: boolean) {}
  toggleColGroup(idx: number) { const g = this.sheets[this.currentSheetIdx]?.colGroups?.[idx]; if (g) g.collapsed = !g.collapsed; if (this.cdr) this.cdr.markForCheck(); }
  toggleFilter() { this.filterActive = !this.filterActive; if (this.cdr) this.cdr.markForCheck(); }
  toggleFooterMenu(m: string) { this.activeFooterMenu = this.activeFooterMenu === m ? null : m; }
  toggleFullScreen() { if (!document.fullscreenElement) document.documentElement.requestFullscreen().catch(() => {}); else document.exitFullscreen().catch(() => {}); }
  toggleGridlines() { const s = this.sheets[this.currentSheetIdx]; if (s) s.hideGridlines = !s.hideGridlines; }
  toggleLockSheet() {}
  toggleMenu(m: string) { this.activeMenu = this.activeMenu === m ? null : m; }
  togglePalette(p: string) { this.activePalette = this.activePalette === p ? null : p; }
  toggleRowGroup(idx: number) { const g = this.sheets[this.currentSheetIdx]?.rowGroups?.[idx]; if (g) g.collapsed = !g.collapsed; if (this.cdr) this.cdr.markForCheck(); }
  toggleSheetGridlines(idx: number) { const s = this.sheets[idx]; if (s) s.hideGridlines = !s.hideGridlines; }
  toggleStar() { this.isStarred = !this.isStarred; }
  toggleViewSetting(key: string) { (this as any)[key] = !(this as any)[key]; }
  toggleWrap() { this.forEachSelectedCell((r, c) => { const k = `${r},${c}`; if (!this.formats[k]) this.formats[k] = {} as any; (this.formats[k] as any).wrap = !(this.formats[k] as any).wrap; }); this.onCellChange(); }
  translateLoading = false;
  translateModalOpen = false;
  translateSheet() {}
  translateSourceText = '';
  translateTargetLang = 'es';
  translateTargetText = '';
  trashDoc() {}
  triggerCopy() { this.copyCell(); }
  triggerImageInsert() {}
  triggerRename() {}
  ungroupCol() {}
  ungroupRow() {}
  unhideAllSheets() { this.sheets.forEach((s: any) => s.hidden = false); }
  unhideCols() {}
  unhideRows() {}
  unhideSheetAndSwitch(idx: number) { this.sheets[idx].hidden = false; this.switchSheet(idx); }
  unmerge() {}
  viewForm() {}
  zoomIn() { this.zoomLevel = Math.min(200, this.zoomLevel + 10); }
  zoomOut() { this.zoomLevel = Math.max(50, this.zoomLevel - 10); }
  onInlineSearch_dummy() {}
"""

# Find the ngOnDestroy method we added and insert before it
target = '  // ---- ngOnDestroy ----'
idx = content.rfind(target)
if idx == -1:
    print("ERROR: could not find ngOnDestroy anchor")
else:
    content = content[:idx] + missing_members + '\n' + content[idx:]
    open(filepath, 'w', encoding='utf-8').write(content)
    print(f"Done. Lines: {len(content.splitlines())}")
