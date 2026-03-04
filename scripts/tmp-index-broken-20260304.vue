<template>
  <div class="excel-editor-page">
    <section class="hero-card">
      <div>
        <p class="eyebrow">???????</p>
        <h1>?? Excel ??</h1>
        <p class="hero-desc">
          ??????????????????????????????????????????????
        </p>
      </div>
      <div class="hero-actions">
        <el-button @click="goToFileCenter">??????</el-button>
        <el-button :loading="loadingServerFile" :disabled="!currentFileName" @click="reloadWorkbook">????</el-button>
        <el-button type="success" :loading="saving" :disabled="!hasWorkbook" @click="saveToServer">????</el-button>
        <el-button :disabled="!currentFileName" @click="downloadServerWorkbook">??????</el-button>
      </div>
    </section>

    <section class="summary-grid">
      <article class="summary-card">
        <span class="summary-label">????</span>
        <strong>{{ currentFileName || '???' }}</strong>
        <span>{{ workbookName }}</span>
      </article>
      <article class="summary-card">
        <span class="summary-label">???</span>
        <strong>{{ currentSheet ? currentSheet.name : '???' }}</strong>
        <span>{{ sheets.length }} ????</span>
      </article>
      <article class="summary-card">
        <span class="summary-label">????</span>
        <strong>{{ isDirty ? `??????${dirtyCount}` : '???' }}</strong>
        <span>{{ isDirty ? '????????????????' : '?????????????' }}</span>
      </article>
    </section>

    <section class="editor-shell" v-loading="pageLoading">
      <template v-if="hasWorkbook">
        <div class="toolbar">
          <div class="toolbar-left">
            <el-tag type="info" effect="light">?????{{ currentFileName }}</el-tag>
            <el-tag :type="isDirty ? 'warning' : 'success'" effect="light">{{ isDirty ? '???' : '???' }}</el-tag>
            <el-tag v-if="lockStatusLabel" :type="lockState.self ? 'success' : 'warning'" effect="light">{{ lockStatusLabel }}</el-tag>
          </div>
          <div class="toolbar-right">
            <span class="selection-tip">{{ selectionLabel }}</span>
            <el-button size="small" :disabled="!canUndo" @click="undoLastAction">??</el-button>
            <el-button size="small" :disabled="!canRedo" @click="redoLastAction">??</el-button>
            <el-button size="small" :disabled="!activeSelection" @click="copySelection(false)">??</el-button>
            <el-button size="small" :disabled="!activeSelection" @click="copySelection(true)">??</el-button>
            <el-button size="small" :disabled="!activeSelection" @click="pasteToSelection">??</el-button>
            <el-button size="small" @click="openFindReplaceDialog">????</el-button>
            <el-button size="small" :disabled="!canMergeSelection" @click="mergeSelectedCells">?????</el-button>
            <el-button size="small" :disabled="!canUnmergeSelection" @click="unmergeSelectedCells">????</el-button>
            <el-button size="small" :disabled="!activeSelection" @click="insertRowAbove">????</el-button>
            <el-button size="small" :disabled="!activeSelection" @click="insertRowBelow">????</el-button>
            <el-button size="small" :disabled="!activeSelection" @click="deleteSelectedRows">???</el-button>
            <el-button size="small" :disabled="!activeSelection" @click="insertColumnLeft">????</el-button>
            <el-button size="small" :disabled="!activeSelection" @click="insertColumnRight">????</el-button>
            <el-button size="small" :disabled="!activeSelection" @click="deleteSelectedColumns">???</el-button>
          </div>
        </div>

        <div class="style-toolbar">
          <span class="style-toolbar-title">??</span>
          <el-button size="small" :disabled="!activeSelection" @click="toggleBold"><strong>B</strong></el-button>
          <el-button size="small" :disabled="!activeSelection" @click="toggleItalic"><em>I</em></el-button>
          <el-button size="small" :disabled="!activeSelection" @click="toggleUnderline"><span class="underline-btn">U</span></el-button>
          <el-select
            class="style-select"
            size="small"
            :model-value="fontSizePicker"
            :disabled="!activeSelection"
            placeholder="??"
            @change="applyFontSize"
          >
            <el-option
              v-for="size in FONT_SIZE_OPTIONS"
              :key="`font-size-${size}`"
              :label="`${size}pt`"
              :value="`${size}pt`"
            />
          </el-select>
          <el-button size="small" :disabled="!activeSelection" @click="applyHorizontalAlign('left')">???</el-button>
          <el-button size="small" :disabled="!activeSelection" @click="applyHorizontalAlign('center')">??</el-button>
          <el-button size="small" :disabled="!activeSelection" @click="applyHorizontalAlign('right')">???</el-button>
          <el-button size="small" :disabled="!activeSelection" @click="applyVerticalAlign('top')">???</el-button>
          <el-button size="small" :disabled="!activeSelection" @click="applyVerticalAlign('middle')">???</el-button>
          <el-button size="small" :disabled="!activeSelection" @click="applyVerticalAlign('bottom')">???</el-button>
          <el-color-picker
            class="style-color-picker"
            size="small"
            v-model="textColorPicker"
            :disabled="!activeSelection"
            :show-alpha="false"
            @active-change="handleTextColorModelValueChange"
            @update:model-value="handleTextColorModelValueChange"
            @change="handleTextColorChange"
          />
          <el-button size="small" :disabled="!activeSelection" @click="clearTextColor">????</el-button>
          <el-color-picker
            class="style-color-picker"
            size="small"
            v-model="fillColorPicker"
            :disabled="!activeSelection"
            :show-alpha="false"
            @active-change="handleFillColorModelValueChange"
            @update:model-value="handleFillColorModelValueChange"
            @change="handleFillColorChange"
          />
          <el-button size="small" :disabled="!activeSelection" @click="clearFillColor">????</el-button>
          <el-button
            size="small"
            :type="formatPainterActive ? 'primary' : 'default'"
            :disabled="!canUseFormatPainter"
            @click="toggleFormatPainter"
          >
            ???
          </el-button>
          <span class="style-toolbar-title">??</span>
          <el-button size="small" :disabled="!activeSelection" @click="applyBorderToSelection('all')">????</el-button>
          <el-button size="small" :disabled="!activeSelection" @click="applyBorderToSelection('outer')">???</el-button>
          <el-button size="small" :disabled="!activeSelection" @click="applyBorderToSelection('clear')">????</el-button>
          <span class="style-toolbar-title">??</span>
          <el-button size="small" :disabled="!activeSelection" @click="formatSelectionAsPercent">???</el-button>
          <el-button size="small" :disabled="!activeSelection" @click="formatSelectionAsDate">??</el-button>
          <el-button size="small" :disabled="!activeSelection" @click="increaseSelectionDecimals">????</el-button>
          <el-button size="small" :disabled="!activeSelection" @click="decreaseSelectionDecimals">????</el-button>
          <span class="style-toolbar-title">??</span>
          <el-button size="small" :disabled="!activeSelection" @click="sortSelection('asc')">??</el-button>
          <el-button size="small" :disabled="!activeSelection" @click="sortSelection('desc')">??</el-button>
          <el-switch v-model="sortHasHeader" size="small" active-text="???" inactive-text="???" />
          <el-button size="small" :disabled="!activeSelection" @click="openSortDialog">????</el-button>
          <span class="style-toolbar-title">??</span>
          <el-button size="small" :disabled="!activeSelection" @click="toggleAutoFilter">????</el-button>
          <el-button size="small" :disabled="!hasActiveFilter" @click="clearAutoFilter">????</el-button>
          <el-button size="small" :disabled="!canOpenColumnFilter" @click="openColumnFilterDialogBySelection">???</el-button>
          <span class="style-toolbar-title">??</span>
          <el-button size="small" :disabled="!currentSheet" @click="freezeFirstRow">????</el-button>
          <el-button size="small" :disabled="!currentSheet" @click="freezeFirstColumn">????</el-button>
          <el-button size="small" :disabled="!activeSelection" @click="freezeBySelection">?????</el-button>
          <el-button size="small" :disabled="!currentSheet" @click="openFreezeDialog">?????</el-button>
          <el-button size="small" :disabled="!currentSheet" @click="clearFreezePane">????</el-button>
          <el-button size="small" @click="shortcutDialogVisible = true">???</el-button>
        </div>

        <div class="formula-toolbar">
          <span class="formula-label">{{ activeCellLabel }}</span>
          <el-input
            v-model="formulaInput"
            class="formula-input"
            clearable
            :disabled="!activePrimaryCell"
            placeholder="???????? = ???"
            @focus="formulaBarFocused = true"
            @blur="formulaBarFocused = false"
            @keydown.enter.prevent="applyFormulaBarInput(false)"
          />
          <el-button size="small" :disabled="!activePrimaryCell" @click="applyFormulaBarInput(false)">??</el-button>
          <el-button size="small" :disabled="!activeSelection" @click="applyFormulaBarInput(true)">????</el-button>
          <span class="formula-preview">{{ formulaPreview }}</span>
        </div>

        <el-tabs v-model="activeSheetName" class="sheet-tabs">
          <el-tab-pane
            v-for="sheet in sheets"
            :key="sheet.name"
            :label="sheet.name"
            :name="sheet.name"
          />
        </el-tabs>
        <div class="sheet-actions">
          <el-button size="small" @click="createSheet">?????</el-button>
          <el-button size="small" :disabled="!currentSheet" @click="renameCurrentSheet">???</el-button>
          <el-button size="small" :disabled="!currentSheet" @click="copyCurrentSheet">?????</el-button>
          <el-button size="small" :disabled="!canDeleteCurrentSheet" @click="deleteCurrentSheet">?????</el-button>
          <el-button size="small" :disabled="!currentSheet || !canMoveCurrentSheetLeft" @click="moveCurrentSheetLeft">??</el-button>
          <el-button size="small" :disabled="!currentSheet || !canMoveCurrentSheetRight" @click="moveCurrentSheetRight">??</el-button>
        </div>

        <div v-if="currentSheet" class="grid-wrap">
          <div class="sheet-stage" :style="buildSheetStageStyle(currentSheet)">
            <div v-if="hasBackgroundImages(currentSheet)" class="sheet-image-layer sheet-image-layer-background">
              <div
                v-for="(image, imageIndex) in listBackgroundImages(currentSheet)"
                :key="buildImageKey(image, imageIndex)"
                class="sheet-image-item background"
                :style="buildImageStyle(image)"
              >
                <img class="sheet-image" :src="image.src" :alt="image.description || '????'" draggable="false" />
              </div>
            </div>

            <table class="sheet-grid">
              <thead>
                <tr>
                  <th class="corner-header"></th>
                  <th
                    v-for="columnIndex in currentSheet.maxColumnCount"
                    :key="`header-col-${columnIndex}`"
                    class="column-header"
                    :style="buildColumnHeaderStyle(currentSheet, columnIndex - 1)"
                    @contextmenu.prevent="openContextMenu($event, { rowIndex: 0, colIndex: columnIndex - 1 })"
                  >
                    <span class="column-label">{{ toColumnLabel(columnIndex - 1) }}</span>
                    <span
                      class="resize-handle col-resize-handle"
                      @mousedown="startColumnResize(columnIndex - 1, $event)"
                    ></span>
                  </th>
                </tr>
              </thead>
              <tbody>
                <tr v-for="row in visibleRows" :key="`row-${row.rowIndex}`">
                  <td class="row-index" :style="buildRowIndexStyle(row)" @contextmenu.prevent="openContextMenu($event, { rowIndex: row.rowIndex, colIndex: 0 })">
                    {{ row.rowIndex + 1 }}
                    <span
                      class="resize-handle row-resize-handle"
                      @mousedown="startRowResize(row.rowIndex, $event)"
                    ></span>
                  </td>
                  <td
                    v-for="cell in row.cells"
                    :key="`cell-${row.rowIndex}-${cell.colIndex}`"
                    class="sheet-cell"
                    :class="buildCellClass(cell)"
                    :colspan="cell.colSpan"
                    :rowspan="cell.rowSpan"
                    :style="buildCellTdStyle(cell)"
                    :data-row-index="cell.rowIndex"
                    :data-col-index="cell.colIndex"
                    @click="handleCellSelection(cell, $event)"
                    @contextmenu.prevent="openContextMenu($event, cell)"
                  >
                    <template v-if="hasValidationOptions(cell)">
                      <select
                        v-model="cell.value"
                        class="cell-input cell-select"
                        :class="{ dirty: cell.dirty, formula: cell.formula }"
                        :style="buildCellInputStyle(cell)"
                        @focus="handleCellFocus(cell)"
                        @blur="handleCellBlur(cell)"
                        @change="handleCellInput(cell)"
                      >
                        <option v-if="cell.validationAllowBlank" value=""></option>
                        <option
                          v-if="shouldKeepCurrentValidationValue(cell)"
                          :value="cell.value"
                        >
                          {{ cell.value }}
                        </option>
                        <option
                          v-for="option in cell.validationOptions"
                          :key="`list-${cell.rowIndex}-${cell.colIndex}-${option}`"
                          :value="option"
                        >
                          {{ option }}
                        </option>
                      </select>
                    </template>
                    <textarea
                      v-else
                      v-model="cell.value"
                      class="cell-input"
                      :class="{ dirty: cell.dirty, formula: cell.formula }"
                      :style="buildCellInputStyle(cell)"
                      :maxlength="MAX_CELL_TEXT_LENGTH"
                      :wrap="cell.style?.whiteSpace === 'pre-wrap' ? 'soft' : 'off'"
                      spellcheck="false"
                      @focus="handleCellFocus(cell)"
                      @blur="handleCellBlur(cell)"
                      @input="handleCellInput(cell)"
                    />
                    <button
                      v-if="isFilterHeaderCell(cell)"
                      type="button"
                      class="filter-trigger"
                      title="???????"
                      @click.stop="openColumnFilterDialog(cell.colIndex)"
                    >
                      ?
                    </button>
                    <span
                      v-if="isFillHandleCell(cell)"
                      class="fill-handle"
                      title="????"
                      @mousedown.stop.prevent="startFillHandleDrag($event)"
                    ></span>
                  </td>
                </tr>
              </tbody>
            </table>

            <div v-if="hasAnchoredImages(currentSheet)" class="sheet-image-layer sheet-image-layer-foreground">
              <div
                v-for="(image, imageIndex) in listAnchoredImages(currentSheet)"
                :key="buildImageKey(image, imageIndex)"
                class="sheet-image-item anchored"
                :style="buildImageStyle(image)"
              >
                <img class="sheet-image" :src="image.src" :alt="image.description || '????'" draggable="false" />
              </div>
            </div>
          </div>
        </div>
      </template>

      <el-empty v-else description="????????????????????????? Excel?" />
    </section>

    <el-dialog v-model="findDialogVisible" title="?????" width="460px" destroy-on-close>
      <el-form label-width="92px">
        <el-form-item label="????">
          <el-input v-model="findForm.findText" placeholder="??????????" />
        </el-form-item>
        <el-form-item label="???">
          <el-input v-model="findForm.replaceText" placeholder="??" />
        </el-form-item>
        <el-form-item label="??">
          <el-radio-group v-model="findForm.scope">
            <el-radio label="sheet">?????</el-radio>
            <el-radio label="workbook">????</el-radio>
          </el-radio-group>
        </el-form-item>
        <el-form-item label="??">
          <div class="find-options">
            <el-checkbox v-model="findForm.matchCase">?????</el-checkbox>
            <el-checkbox v-model="findForm.wholeWord">????</el-checkbox>
            <el-checkbox v-model="findForm.selectionOnly" :disabled="!activeSelection || findForm.scope === 'workbook'">
              ?????
            </el-checkbox>
          </div>
        </el-form-item>
        <el-form-item label="??">
          <el-tag type="info" effect="light">???? {{ findPreviewCount }} ?</el-tag>
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="findDialogVisible = false">??</el-button>
        <el-button @click="previewFindMatches">????</el-button>
        <el-button @click="findNext">?????</el-button>
        <el-button type="primary" :disabled="findPreviewCount <= 0" @click="replaceAll">????</el-button>
      </template>
    </el-dialog>

    <el-dialog v-model="sortDialogVisible" title="????" width="520px" destroy-on-close>
      <el-form label-width="88px">
        <el-form-item label="????">
          <el-input :model-value="sortDialogRangeLabel" readonly />
        </el-form-item>
        <el-form-item label="??">
          <el-switch v-model="sortDialogHasHeader" active-text="?????" inactive-text="???" />
        </el-form-item>
      </el-form>
      <div class="sort-rule-list">
        <div v-for="(rule, ruleIndex) in sortRules" :key="`sort-rule-${ruleIndex}`" class="sort-rule-item">
          <el-select v-model="rule.colIndex" size="small" class="sort-rule-column">
            <el-option
              v-for="option in sortDialogColumnOptions"
              :key="`sort-col-${option.colIndex}`"
              :label="option.label"
              :value="option.colIndex"
            />
          </el-select>
          <el-select v-model="rule.direction" size="small" class="sort-rule-direction">
            <el-option label="??" value="asc" />
            <el-option label="??" value="desc" />
          </el-select>
          <el-button size="small" :disabled="sortRules.length <= 1" @click="removeSortRule(ruleIndex)">??</el-button>
        </div>
      </div>
      <template #footer>
        <el-button @click="addSortRule">?????</el-button>
        <el-button @click="sortDialogVisible = false">??</el-button>
        <el-button type="primary" @click="applySortRules">????</el-button>
      </template>
    </el-dialog>

    <el-dialog v-model="columnFilterDialogVisible" title="???" width="520px" destroy-on-close>
      <el-form label-width="88px">
        <el-form-item label="?">
          <el-select v-model="filterDialogColIndex" size="small" style="width: 220px" @change="resetFilterCandidateValues">
            <el-option
              v-for="option in filterColumnOptions"
              :key="`filter-col-${option.colIndex}`"
              :label="option.label"
              :value="option.colIndex"
            />
          </el-select>
        </el-form-item>
        <el-form-item label="???">
          <el-input v-model="filterKeyword" clearable placeholder="??????????" />
        </el-form-item>
        <el-form-item label="???">
          <div class="filter-value-box">
            <el-checkbox-group v-model="filterSelectedValues">
              <el-checkbox
                v-for="item in filteredCandidateValues"
                :key="`filter-value-${item.value}`"
                :label="item.value"
              >
                {{ item.label }}
              </el-checkbox>
            </el-checkbox-group>
          </div>
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="selectAllFilterValues">??</el-button>
        <el-button @click="clearFilterValuesSelection">???</el-button>
        <el-button @click="columnFilterDialogVisible = false">??</el-button>
        <el-button type="primary" @click="applyColumnFilterValues">????</el-button>
      </template>
    </el-dialog>

    <el-dialog v-model="freezeDialogVisible" title="???????" width="420px" destroy-on-close>
      <el-form label-width="96px">
        <el-form-item label="????">
          <el-input-number
            v-model="freezeForm.rows"
            :min="0"
            :max="Math.max((currentSheet?.rowCount || 1) - 1, 0)"
            controls-position="right"
          />
        </el-form-item>
        <el-form-item label="????">
          <el-input-number
            v-model="freezeForm.cols"
            :min="0"
            :max="Math.max((currentSheet?.maxColumnCount || 1) - 1, 0)"
            controls-position="right"
          />
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="freezeDialogVisible = false">??</el-button>
        <el-button type="primary" @click="applyCustomFreezePane">??</el-button>
      </template>
    </el-dialog>

    <el-dialog v-model="shortcutDialogVisible" title="Excel ???" width="560px" destroy-on-close>
      <div class="shortcut-grid">
        <div v-for="item in shortcutList" :key="item.key" class="shortcut-item">
          <code>{{ item.key }}</code>
          <span>{{ item.desc }}</span>
        </div>
      </div>
      <template #footer>
        <el-button type="primary" @click="shortcutDialogVisible = false">????</el-button>
      </template>
    </el-dialog>

    <div
      v-if="contextMenuState.visible"
      class="sheet-context-menu"
      :style="{ left: `${contextMenuState.x}px`, top: `${contextMenuState.y}px` }"
    >
      <button type="button" @click="insertRowAbove">????</button>
      <button type="button" @click="insertRowBelow">????</button>
      <button type="button" @click="deleteSelectedRows">???</button>
      <button type="button" @click="insertColumnLeft">????</button>
      <button type="button" @click="insertColumnRight">????</button>
      <button type="button" @click="deleteSelectedColumns">???</button>
      <button type="button" @click="copySelection(false)">??</button>
      <button type="button" @click="copySelection(true)">??</button>
      <button type="button" @click="pasteToSelection">??</button>
    </div>
  </div>
</template>
                    <textarea
                      v-else
                      v-model="cell.value"
                      class="cell-input"
                      :class="{ dirty: cell.dirty, formula: cell.formula }"
                      :style="buildCellInputStyle(cell)"
                      :maxlength="MAX_CELL_TEXT_LENGTH"
                      :wrap="cell.style?.whiteSpace === 'pre-wrap' ? 'soft' : 'off'"
                      spellcheck="false"
                      @focus="handleCellFocus(cell)"
                      @blur="handleCellBlur(cell)"
                      @input="handleCellInput(cell)"
                    />
                    <button
                      v-if="isFilterHeaderCell(cell)"
                      type="button"
                      class="filter-trigger"
                      title="鎵撳紑鍒楃瓫閫夐潰鏉?
                      @click.stop="openColumnFilterDialog(cell.colIndex)"
                    >
                      鈻?                    </button>
                    <span
                      v-if="isFillHandleCell(cell)"
                      class="fill-handle"
                      title="鎷栨嫿濉厖"
                      @mousedown.stop.prevent="startFillHandleDrag($event)"
                    ></span>
                  </td>
                </tr>
              </tbody>
            </table>

            <div v-if="hasAnchoredImages(currentSheet)" class="sheet-image-layer sheet-image-layer-foreground">
              <div
                v-for="(image, imageIndex) in listAnchoredImages(currentSheet)"
                :key="buildImageKey(image, imageIndex)"
                class="sheet-image-item anchored"
                :style="buildImageStyle(image)"
              >
                <img class="sheet-image" :src="image.src" :alt="image.description || '宓屽叆鍥剧墖'" draggable="false" />
              </div>
            </div>
          </div>
        </div>
      </template>

      <el-empty v-else description="杩樻病鏈夊姞杞藉彲缂栬緫鏂囦欢锛岃鍏堝湪鏂囦欢绠＄悊涓€夋嫨宸蹭笂浼犵殑 Excel銆? />
    </section>
    <el-dialog v-model="findDialogVisible" title="鏌ユ壘涓庢浛鎹? width="460px" destroy-on-close>
const workbookName = ref('???')
        <el-form-item label="鏌ユ壘鍐呭">
          <el-input v-model="findForm.findText" placeholder="璇疯緭鍏ラ渶瑕佹煡鎵剧殑鏂囨湰" />
        </el-form-item>
        <el-form-item label="鏇挎崲涓?>
          <el-input v-model="findForm.replaceText" placeholder="鍙€? />
        </el-form-item>
        <el-form-item label="鑼冨洿">
          <el-radio-group v-model="findForm.scope">
            <el-radio label="sheet">褰撳墠宸ヤ綔琛?/el-radio>
            <el-radio label="workbook">鍏ㄥ伐浣滅翱</el-radio>
          </el-radio-group>
        </el-form-item>
        <el-form-item label="鍖归厤">
          <div class="find-options">
            <el-checkbox v-model="findForm.matchCase">鍖哄垎澶у皬鍐?/el-checkbox>
            <el-checkbox v-model="findForm.wholeWord">鏁磋瘝鍖归厤</el-checkbox>
            <el-checkbox v-model="findForm.selectionOnly" :disabled="!activeSelection || findForm.scope === 'workbook'">
              浠呭綋鍓嶉€夊尯
            </el-checkbox>
          </div>
        </el-form-item>
        <el-form-item label="棰勮">
          <el-tag type="info" effect="light">棰勮鍖归厤 {{ findPreviewCount }} 澶?/el-tag>
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="findDialogVisible = false">鍏抽棴</el-button>
        <el-button @click="previewFindMatches">棰勮鏁伴噺</el-button>
        <el-button @click="findNext">鏌ユ壘涓嬩竴涓?/el-button>
        <el-button type="primary" :disabled="findPreviewCount <= 0" @click="replaceAll">鍏ㄩ儴鏇挎崲</el-button>
      </template>
    </el-dialog>
    <el-dialog v-model="sortDialogVisible" title="澶氬垪鎺掑簭" width="520px" destroy-on-close>
      <el-form label-width="88px">
        <el-form-item label="鏁版嵁鍖哄煙">
          <el-input :model-value="sortDialogRangeLabel" readonly />
        </el-form-item>
        <el-form-item label="琛ㄥご">
          <el-switch v-model="sortDialogHasHeader" active-text="棣栬涓鸿〃澶? inactive-text="鏃犺〃澶? />
        </el-form-item>
      </el-form>
      <div class="sort-rule-list">
        <div v-for="(rule, ruleIndex) in sortRules" :key="`sort-rule-${ruleIndex}`" class="sort-rule-item">
          <el-select v-model="rule.colIndex" size="small" class="sort-rule-column">
            <el-option
              v-for="option in sortDialogColumnOptions"
              :key="`sort-col-${option.colIndex}`"
              :label="option.label"
              :value="option.colIndex"
            />
          </el-select>
          <el-select v-model="rule.direction" size="small" class="sort-rule-direction">
            <el-option label="鍗囧簭" value="asc" />
            <el-option label="闄嶅簭" value="desc" />
          </el-select>
          <el-button size="small" :disabled="sortRules.length <= 1" @click="removeSortRule(ruleIndex)">鍒犻櫎</el-button>
        </div>
      </div>
      <template #footer>
        <el-button @click="addSortRule">娣诲姞鎺掑簭鍒?/el-button>
        <el-button @click="sortDialogVisible = false">鍙栨秷</el-button>
        <el-button type="primary" @click="applySortRules">寮€濮嬫帓搴?/el-button>
      </template>
    </el-dialog>
    <el-dialog v-model="columnFilterDialogVisible" title="鍒楃瓫閫? width="520px" destroy-on-close>
      <el-form label-width="88px">
        <el-form-item label="鍒?>
          <el-select v-model="filterDialogColIndex" size="small" style="width: 220px" @change="resetFilterCandidateValues">
            <el-option
              v-for="option in filterColumnOptions"
  { key: 'Ctrl+S', desc: '???????' },
              :label="option.label"
              :value="option.colIndex"
  { key: 'Ctrl+Shift+ArrowDown', desc: '???????????????????' },
          </el-select>
  { key: 'Ctrl+Shift+L', desc: '?????????' },
        <el-form-item label="鍏抽敭瀛?>
  { key: 'Alt+Shift+ArrowUp / Alt+Shift+ArrowDown', desc: '?????? / ??' },
        </el-form-item>
  { key: 'F2', desc: '?????????' },
          <div class="filter-value-box">
            <el-checkbox-group v-model="filterSelectedValues">
              <el-checkbox
                v-for="item in filteredCandidateValues"
                :key="`filter-value-${item.value}`"
                :label="item.value"
              >
                {{ item.label }}
              </el-checkbox>
            </el-checkbox-group>
          </div>
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="selectAllFilterValues">鍏ㄩ€?/el-button>
        <el-button @click="clearFilterValuesSelection">鍏ㄤ笉閫?/el-button>
        <el-button @click="columnFilterDialogVisible = false">鍙栨秷</el-button>
        <el-button type="primary" @click="applyColumnFilterValues">搴旂敤绛涢€?/el-button>
      </template>
    </el-dialog>
    <el-dialog v-model="freezeDialogVisible" title="鑷畾涔夊喕缁撶獥鏍? width="420px" destroy-on-close>
      <el-form label-width="96px">
        <el-form-item label="鍐荤粨琛屾暟">
          <el-input-number
            v-model="freezeForm.rows"
            :min="0"
            :max="Math.max((currentSheet?.rowCount || 1) - 1, 0)"
            controls-position="right"
          />
        </el-form-item>
        <el-form-item label="鍐荤粨鍒楁暟">
          <el-input-number
            v-model="freezeForm.cols"
            :min="0"
            :max="Math.max((currentSheet?.maxColumnCount || 1) - 1, 0)"
            controls-position="right"
          />
        </el-form-item>
    return '?'
      <template #footer>
        <el-button @click="freezeDialogVisible = false">鍙栨秷</el-button>
    return cell.value ? `??${cell.value}` : '?'
      </template>
  return cell.displayValue ? `???${cell.displayValue}` : '?????'
    <el-dialog v-model="shortcutDialogVisible" title="Excel 蹇嵎閿? width="560px" destroy-on-close>
      <div class="shortcut-grid">
        <div v-for="item in shortcutList" :key="item.key" class="shortcut-item">
          <code>{{ item.key }}</code>
          <span>{{ item.desc }}</span>
        </div>
      </div>
      <template #footer>
        <el-button type="primary" @click="shortcutDialogVisible = false">鎴戠煡閬撲簡</el-button>
      </template>
    </el-dialog>

    <div
      v-if="contextMenuState.visible"
      class="sheet-context-menu"
      :style="{ left: `${contextMenuState.x}px`, top: `${contextMenuState.y}px` }"
    >
      <button type="button" @click="insertRowAbove">涓婃彃鍏ヨ</button>
      <button type="button" @click="insertRowBelow">涓嬫彃鍏ヨ</button>
      <button type="button" @click="deleteSelectedRows">鍒犻櫎琛?/button>
      <button type="button" @click="insertColumnLeft">宸︽彃鍏ュ垪</button>
      <button type="button" @click="insertColumnRight">鍙虫彃鍏ュ垪</button>
      <button type="button" @click="deleteSelectedColumns">鍒犻櫎鍒?/button>
      <button type="button" @click="copySelection(false)">澶嶅埗</button>
      <button type="button" @click="copySelection(true)">鍓垏</button>
      <button type="button" @click="pasteToSelection">绮樿创</button>
    </div>
  </div>
</template>

<script setup>
import { computed, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import { useRoute, useRouter } from 'vue-router'
import { ElMessage, ElMessageBox } from 'element-plus'
import { saveAs } from 'file-saver'
import {
  acquireExcelEditorLock,
  applyExcelWorkbookChanges,
  getExcelEditorContent,
  getExcelEditorInfo,
  getExcelWorkbookView,
  heartbeatExcelEditorLock,
  releaseExcelEditorLock
} from '@/api/tool/excelEditor'

const DEFAULT_ROW_HEIGHT = 28
const MIN_ROW_HEIGHT = 24
const DEFAULT_COLUMN_WIDTH = 96
const MIN_COLUMN_WIDTH = 72
const DEFAULT_CELL_HEIGHT = 32
const COLUMN_HEADER_HEIGHT = 34
const ROW_INDEX_WIDTH = 56
const MAX_CELL_TEXT_LENGTH = 32767
const LOCK_HEARTBEAT_INTERVAL = 30000
const AUTO_SAVE_INTERVAL = 60000
const MAX_HISTORY_STEPS = 100
const DEFAULT_EMPTY_SHEET_ROW_COUNT = 120
const DEFAULT_EMPTY_SHEET_COLUMN_COUNT = 30
const BORDER_LINE_STYLE = '1px solid #9CA3AF'
const MAX_DECIMAL_SCALE = 8
const EDITABLE_STYLE_KEYS = [
  'textAlign',
  'verticalAlign',
  'color',
  'backgroundColor',
  'fontWeight',
  'fontStyle',
  'textDecoration',
  'fontSize',
  'fontFamily',
  'whiteSpace',
  'borderTop',
  'borderRight',
  'borderBottom',
  'borderLeft',
  'numberFormat'
]
const FONT_SIZE_OPTIONS = [10, 11, 12, 14, 16, 18, 20, 24, 28, 32]

const route = useRoute()
const router = useRouter()

const workbookName = ref('鏈姞杞?)
const sheets = ref([])
const activeSheetName = ref('')
const currentFileName = ref('')
const currentWorkbookVersion = ref('')
const pageLoading = ref(false)
const loadingServerFile = ref(false)
const saving = ref(false)
const isDirty = ref(false)
const dirtyCount = ref(0)
const revertingRoute = ref(false)
const selection = ref(null)
const showedCellLimitHint = ref(false)
const columnResizeState = ref(null)
const rowResizeState = ref(null)
const lockState = ref({
  fileName: '',
  ownerUsername: '',
  expiresAt: 0,
  self: false
})
const lockHeartbeatTimer = ref(null)
const autoSaveTimer = ref(null)
const lockWarningShown = ref(false)
const dirtyRefreshRafId = ref(0)
const undoStack = ref([])
const redoStack = ref([])
const applyingHistory = ref(false)
const activeEditCellKey = ref('')
const clipboardPayload = ref(null)
const pendingStructures = ref([])
const contextMenuState = ref({
  visible: false,
  x: 0,
  y: 0,
  rowIndex: 0,
  colIndex: 0
})
const findDialogVisible = ref(false)
const findForm = ref({
  findText: '',
  replaceText: '',
  scope: 'sheet',
  matchCase: false,
  wholeWord: false,
  selectionOnly: false
})
const formatPainterStyle = ref(null)
const formatPainterActive = ref(false)
const textColorPicker = ref('#111827')
const fillColorPicker = ref('#FFFFFF')
const fontSizePicker = ref('12pt')
const formulaInput = ref('')
const formulaBarFocused = ref(false)
const sortHasHeader = ref(true)
const sortDialogVisible = ref(false)
const sortDialogRange = ref(null)
const sortDialogHasHeader = ref(true)
const sortRules = ref([])
const columnFilterDialogVisible = ref(false)
const filterDialogColIndex = ref(null)
const filterKeyword = ref('')
const filterSelectedValues = ref([])
const freezeDialogVisible = ref(false)
const freezeForm = ref({
  rows: 0,
  cols: 0
})
const fillDragState = ref(null)
const shortcutDialogVisible = ref(false)
const shortcutList = Object.freeze([
  { key: 'Ctrl+S', desc: '淇濆瓨褰撳墠宸ヤ綔绨? },
  { key: 'Ctrl+Z / Ctrl+Y', desc: '鎾ら攢 / 閲嶅仛' },
  { key: 'Ctrl+C / Ctrl+X / Ctrl+V', desc: '澶嶅埗 / 鍓垏 / 绮樿创' },
  { key: 'Ctrl+Shift+鈫?, desc: '浠庡綋鍓嶅崟鍏冩牸蹇€熼€変腑鏈垪鍚戜笅鎵€鏈夊崟鍏冩牸' },
  { key: '鎷栨嫿濉厖鏌?, desc: '鍚戜笅/鍚戝彸鎷栨嫿濉厖搴忓垪鎴栧鍒跺叕寮? },
  { key: 'Ctrl+Shift+L', desc: '寮€鍚垨鍏抽棴鑷姩绛涢€? },
  { key: 'Alt+S', desc: '鎵撳紑澶氬垪鎺掑簭闈㈡澘' },
  { key: 'Alt+Shift+鈫?/ Alt+Shift+鈫?, desc: '褰撳墠閫夊尯鍗囧簭 / 闄嶅簭' },
  { key: 'Alt+鈫?, desc: '鎵撳紑褰撳墠鍗曞厓鏍间笅鎷夊垪琛? },
  { key: 'F2', desc: '鑱氱劍褰撳墠鍗曞厓鏍肩紪杈? },
  { key: 'Ctrl+Enter', desc: '濉厖褰撳墠閫夊尯' },
  { key: 'Ctrl+F / Ctrl+H', desc: '鏌ユ壘 / 鏇挎崲' }
])

const currentSheet = computed(() => sheets.value.find((sheet) => sheet.name === activeSheetName.value) || null)
const hasWorkbook = computed(() => sheets.value.length > 0)
const canUndo = computed(() => undoStack.value.length > 0)
const canRedo = computed(() => redoStack.value.length > 0)
const activeSelection = computed(() => {
  if (!selection.value || selection.value.sheetName !== activeSheetName.value) {
    return null
  }
  return selection.value
})
const selectionLabel = computed(() => {
  if (!activeSelection.value) {
    return '鐐瑰嚮鍗曞厓鏍奸€夋嫨锛屾寜浣?Shift 鍙墿灞曢€夊尯'
  }
  return `閫夊尯 ${formatSelectionRange(activeSelection.value)}`
})
const activePrimaryCell = computed(() => {
  const sheet = currentSheet.value
  const range = activeSelection.value
  if (!sheet || !range) {
    return null
  }
  return findCellByCoordinate(sheet, range.startRow, range.startCol)
})
const activeCellLabel = computed(() => {
  const range = activeSelection.value
  if (!range) {
    return '--'
  }
  return `${toColumnLabel(range.startCol)}${range.startRow + 1}`
})
const formulaPreview = computed(() => {
  const cell = activePrimaryCell.value
    await ElMessageBox.confirm('?????????????????????????????', '????', {
    return '绌?
  }
  if (!cell.formula) {
    return cell.value ? `鍊硷細${cell.value}` : '绌?
  }
  return cell.displayValue ? `缁撴灉锛?{cell.displayValue}` : '缁撴灉锛堢┖锛?
})
const visibleRows = computed(() => {
  const sheet = currentSheet.value
  if (!sheet || !Array.isArray(sheet.rows)) {
    return []
  }
  const filterConfig = normalizeFilterConfig(sheet.filterConfig)
  if (!filterConfig.enabled) {
    return sheet.rows
  }
  return sheet.rows.filter((row) => isRowVisibleByFilter(sheet, row.rowIndex, filterConfig))
})
const hasActiveFilter = computed(() => {
  const sheet = currentSheet.value
  if (!sheet) {
    return false
  }
  return normalizeFilterConfig(sheet.filterConfig).enabled
})
const canOpenColumnFilter = computed(() => {
  const sheet = currentSheet.value
  if (!sheet) {
    return false
  }
  const filterConfig = normalizeFilterConfig(sheet.filterConfig)
  if (!filterConfig.enabled) {
    return false
  }
  const range = activeSelection.value
  if (!range) {
    return true
  }
  return isColumnInsideFilterRange(range.startCol, filterConfig)
})
const sortDialogRangeLabel = computed(() => {
  const range = sortDialogRange.value
  if (!range) {
    return ''
  }
  return formatSelectionRange(range)
})
const sortDialogColumnOptions = computed(() => {
  const range = sortDialogRange.value
  if (!range) {
    return []
  }
  const options = []
  for (let colIndex = range.startCol; colIndex <= range.endCol; colIndex += 1) {
    options.push({
      colIndex,
      label: `${toColumnLabel(colIndex)} 鍒梎
    })
  }
  return options
})
const filterColumnOptions = computed(() => {
  const sheet = currentSheet.value
  if (!sheet) {
    return []
  }
  const filterConfig = normalizeFilterConfig(sheet.filterConfig)
  if (!filterConfig.enabled) {
    return []
  }
  const options = []
  for (let colIndex = filterConfig.startCol; colIndex <= filterConfig.endCol; colIndex += 1) {
    options.push({
      colIndex,
      label: `${toColumnLabel(colIndex)} 鍒梎
    })
  }
  return options
})
const filteredCandidateValues = computed(() => {
  const sheet = currentSheet.value
  if (!sheet || filterDialogColIndex.value == null) {
    return []
  }
  const allValues = collectFilterCandidateValues(sheet, Number(filterDialogColIndex.value))
  const keyword = normalizeValue(filterKeyword.value).trim().toLowerCase()
  return allValues.filter((item) => normalizeValue(item.label).toLowerCase().includes(keyword))
})
const canMergeSelection = computed(() => {
  if (!currentSheet.value || !activeSelection.value) {
    return false
  }
  return getSelectionArea(activeSelection.value) > 1
})
const canUnmergeSelection = computed(() => {
  if (!currentSheet.value || !activeSelection.value) {
    return false
  }
  return currentSheet.value.cells.some((cell) => isMergedCell(cell) && rangesIntersect(activeSelection.value, getCellRange(cell)))
})
const lockStatusLabel = computed(() => {
  if (!currentFileName.value || !lockState.value.fileName) {
    return ''
  }
  if (lockState.value.self) {
    return `缂栬緫閿侊細浣狅紙${lockState.value.ownerUsername || '褰撳墠鐢ㄦ埛'}锛塦
  }
  if (lockState.value.ownerUsername) {
    return `琚?${lockState.value.ownerUsername} 閿佸畾`
  }
  return ''
})
const currentSheetIndex = computed(() => sheets.value.findIndex((sheet) => sheet.name === activeSheetName.value))
const canDeleteCurrentSheet = computed(() => sheets.value.length > 1 && currentSheetIndex.value >= 0)
const canMoveCurrentSheetLeft = computed(() => currentSheetIndex.value > 0)
const canMoveCurrentSheetRight = computed(() => currentSheetIndex.value >= 0 && currentSheetIndex.value < sheets.value.length - 1)
const canUseFormatPainter = computed(() => Boolean(activePrimaryCell.value))
const findPreviewCount = computed(() => {
  const query = normalizeValue(findForm.value.findText).trim()
  if (!query) {
    return 0
  }
  const matches = collectFindMatches(query, buildFindOptions())
  return matches.reduce((total, item) => total + Number(item.count || 0), 0)
})

onMounted(async () => {
  window.addEventListener('keydown', handleGlobalKeydown, true)
  window.addEventListener('click', hideContextMenu)
  startAutoSave()
  await initializeFromRoute()
})

onBeforeUnmount(() => {
  window.removeEventListener('keydown', handleGlobalKeydown, true)
  window.removeEventListener('click', hideContextMenu)
  stopFillHandleDrag(false)
  cancelDirtyStateRefresh()
  clearResizeState()
  stopAutoSave()
  stopLockHeartbeat()
  releaseActiveLock(false)
})

watch(
  () => route.query.fileName,
  async (nextValue, previousValue) => {
    if (revertingRoute.value) {
      revertingRoute.value = false
      return
    }
    if (normalizeQueryFileName(nextValue) === normalizeQueryFileName(previousValue)) {
      return
    }
    const confirmed = await confirmReplaceWorkbook()
    if (!confirmed) {
      revertingRoute.value = true
      router.replace({
        path: '/excel-editor/index',
        query: currentFileName.value ? { fileName: currentFileName.value } : {}
      })
      return
    }
    await initializeFromRoute()
  }
)

watch(activeSheetName, () => {
  clearResizeState()
  stopFillHandleDrag(false)
  clearFormatPainterState()
  clearSelection()
})

watch(activePrimaryCell, (cell) => {
  if (formulaBarFocused.value) {
    return
  }
  formulaInput.value = cell ? normalizeValue(cell.value) : ''
})

async function initializeFromRoute() {
  const fileName = normalizeQueryFileName(route.query.fileName)
  if (fileName) {
    await loadWorkbookByFileName(fileName, true)
    return
  }
  try {
    const response = await getExcelEditorInfo()
    const fallbackFileName = response?.data?.currentFileName || ''
    if (!fallbackFileName) {
      clearWorkbook()
      return
    }
    await loadWorkbookByFileName(fallbackFileName, true)
  } catch (error) {
    clearWorkbook()
    ElMessage.error(error?.message || '鍔犺浇鏂囦欢淇℃伅澶辫触')
  }
}

async function confirmReplaceWorkbook() {
  if (!hasWorkbook.value || !isDirty.value) {
    return true
  }
  try {
    await ElMessageBox.confirm('褰撳墠宸ヤ綔绨挎湁鏈繚瀛樹慨鏀癸紝鍒囨崲鏂囦欢浼氫涪澶辫繖浜涙敼鍔紝鏄惁缁х画锛?, '纭鍒囨崲', {
      type: 'warning',
      confirmButtonText: '缁х画',
      cancelButtonText: '鍙栨秷'
    })
    return true
  } catch {
    return false
  }
}

function resetEditingSessionState() {
  undoStack.value = []
  redoStack.value = []
  pendingStructures.value = []
  activeEditCellKey.value = ''
  clipboardPayload.value = null
  formatPainterStyle.value = null
  formatPainterActive.value = false
  formulaInput.value = ''
  formulaBarFocused.value = false
  sortDialogVisible.value = false
  sortDialogRange.value = null
  sortRules.value = []
  columnFilterDialogVisible.value = false
  filterDialogColIndex.value = null
  filterKeyword.value = ''
  filterSelectedValues.value = []
  freezeDialogVisible.value = false
  fillDragState.value = null
  shortcutDialogVisible.value = false
  stopFillHandleDrag(false)
  hideContextMenu()
}

function captureWorkbookSnapshot() {
  return JSON.parse(JSON.stringify({
    sheets: sheets.value,
    activeSheetName: activeSheetName.value,
    selection: selection.value,
    pendingStructures: pendingStructures.value
  }))
}

function restoreWorkbookSnapshot(snapshot) {
  if (!snapshot) {
    return
  }
  applyingHistory.value = true
  sheets.value = Array.isArray(snapshot.sheets) ? snapshot.sheets : []
  activeSheetName.value = snapshot.activeSheetName || sheets.value[0]?.name || ''
  selection.value = snapshot.selection ? { ...snapshot.selection } : null
  pendingStructures.value = Array.isArray(snapshot.pendingStructures) ? snapshot.pendingStructures : []
  for (const sheet of sheets.value) {
    recomputeSheetFormulaDisplays(sheet)
    ensureSheetDirtyMetrics(sheet)
    rebuildSheetRows(sheet)
  }
  refreshDirtyState()
  applyingHistory.value = false
}

function pushUndoSnapshot() {
  if (applyingHistory.value || !hasWorkbook.value) {
    return
  }
  undoStack.value.push(captureWorkbookSnapshot())
  if (undoStack.value.length > MAX_HISTORY_STEPS) {
    undoStack.value.shift()
  }
  redoStack.value = []
}

function undoLastAction() {
  if (!undoStack.value.length) {
    return
  }
  const snapshot = undoStack.value.pop()
  redoStack.value.push(captureWorkbookSnapshot())
  restoreWorkbookSnapshot(snapshot)
}

    ElMessage.warning('???????????')
  if (!redoStack.value.length) {
    return
  }
  const snapshot = redoStack.value.pop()
  undoStack.value.push(captureWorkbookSnapshot())
  restoreWorkbookSnapshot(snapshot)
}

function handleGlobalKeydown(event) {
  if (!hasWorkbook.value) {
    return
  }
  if (event.isComposing || event.key === 'Process') {
    return
  }
  const typing = isTypingTarget(event.target)
  const key = String(event.key || '').toLowerCase()
  const withMeta = event.ctrlKey || event.metaKey
  const allowTypingShortcut = key === 's'
    || key === 'enter'
    || (event.shiftKey && key === 'arrowdown')

  if (withMeta) {
    if (typing && !allowTypingShortcut) {
      return
    }
    if (key === 'z') {
      event.preventDefault()
      if (event.shiftKey) {
        redoLastAction()
      } else {
        undoLastAction()
      }
      return
    }
    if (key === 'y') {
      event.preventDefault()
      redoLastAction()
      return
    }
    if (key === 'c') {
      event.preventDefault()
      copySelection(false)
      return
    }
    if (key === 'x') {
      event.preventDefault()
      copySelection(true)
      return
    }
    if (key === 'v') {
      event.preventDefault()
      pasteToSelection()
      return
    }
    if (key === 'f') {
      event.preventDefault()
      openFindReplaceDialog()
      return
    }
    if (key === 'h') {
      event.preventDefault()
      openFindReplaceDialog()
      return
    }
    if (key === 'enter') {
      event.preventDefault()
      applyFormulaBarInput(true)
      return
    }
    if (event.shiftKey && key === 'l') {
      event.preventDefault()
      toggleAutoFilter()
      return
    }
    if (event.shiftKey && key === 'arrowdown') {
      event.preventDefault()
      selectColumnDownFromActiveCell()
      return
    }
    if (key === 's') {
      event.preventDefault()
      saveToServer()
      return
    }
  }

  if (!withMeta && event.altKey) {
    if (key === 's') {
      event.preventDefault()
      openSortDialog()
      return
    }
    if (event.shiftKey && key === 'arrowup') {
      event.preventDefault()
      sortSelection('asc')
      return
    }
    if (event.shiftKey && key === 'arrowdown') {
      event.preventDefault()
      sortSelection('desc')
      return
    }
    if (key === 'arrowdown') {
      event.preventDefault()
      openValidationDropdownForSelection()
      return
    }
  }

  if (typing) {
    return
  }

  if (event.key === 'F2') {
    event.preventDefault()
    focusActiveCellEditor()
    return
  }
  if (event.key === 'Delete') {
    event.preventDefault()
    clearSelectedCellValues()
    return
  }
  if (event.key === 'ArrowUp') {
    event.preventDefault()
    moveSelectionBy(-1, 0, event.shiftKey)
    return
  }
  if (event.key === 'ArrowDown') {
    event.preventDefault()
    moveSelectionBy(1, 0, event.shiftKey)
    return
  }
  if (event.key === 'ArrowLeft') {
    event.preventDefault()
    moveSelectionBy(0, -1, event.shiftKey)
    return
  }
  if (event.key === 'ArrowRight') {
    event.preventDefault()
    moveSelectionBy(0, 1, event.shiftKey)
    return
  }
  if (event.key === 'Enter') {
    event.preventDefault()
    moveSelectionBy(1, 0, event.shiftKey)
    return
  }
  if (event.key === 'Tab') {
    event.preventDefault()
    moveSelectionBy(0, event.shiftKey ? -1 : 1, event.shiftKey)
  }
}

function isTypingTarget(target) {
  const element = target instanceof HTMLElement ? target : null
  if (!element) {
    return false
  }
  const tagName = element.tagName
  if (tagName === 'INPUT' || tagName === 'TEXTAREA' || tagName === 'SELECT') {
    return true
  }
  return element.isContentEditable
}

function applyFormulaBarInput(fillSelection = false) {
  const sheet = currentSheet.value
  const range = activeSelection.value
  const targetCell = activePrimaryCell.value
  if (!sheet || !targetCell) {
    return
  }
  const nextValue = normalizeValue(formulaInput.value)
  pushUndoSnapshot()
  let applied = 0
  let invalid = 0
  const applyForCell = (cell) => {
    const result = applyCellInputValue(sheet, cell, nextValue, { skipValidation: false })
    if (result.applied) {
      applied += 1
      return
    }
    invalid += 1
  }
  if (fillSelection && range) {
    const touched = new Set()
    for (let rowIndex = range.startRow; rowIndex <= range.endRow; rowIndex += 1) {
      for (let colIndex = range.startCol; colIndex <= range.endCol; colIndex += 1) {
        const cell = findCellByCoordinate(sheet, rowIndex, colIndex)
        if (!cell) {
          continue
        }
        const key = buildCellKey(cell.rowIndex, cell.colIndex)
        if (touched.has(key)) {
          continue
        }
        touched.add(key)
        applyForCell(cell)
      }
    }
  } else {
    applyForCell(targetCell)
  }
  if (!applied) {
    ElMessage.warning('娌℃湁鍙簲鐢ㄥ埌閫夊尯鐨勫唴瀹?)
    return
  }
  recomputeSheetFormulaDisplays(sheet)
  recountSheetCellDirtyCount(sheet)
  scheduleDirtyStateRefresh()
  if (invalid > 0) {
    ElMessage.warning(`鏈?${invalid} 涓崟鍏冩牸鍥犱笅鎷夋牎楠屾湭閫氳繃琚烦杩嘸)
  }
}

function focusActiveCellEditor() {
  const cell = activePrimaryCell.value
  if (!cell) {
    return
  }
  const selector = `td.sheet-cell[data-row-index=\"${cell.rowIndex}\"][data-col-index=\"${cell.colIndex}\"] .cell-input`
  const inputElement = document.querySelector(selector)
  if (!(inputElement instanceof HTMLElement)) {
    return
  }
  inputElement.focus()
  if (typeof inputElement.select === 'function') {
    inputElement.select()
  }
}

function openValidationDropdownForSelection() {
  const cell = activePrimaryCell.value
  if (!cell || !hasValidationOptions(cell)) {
    return
  }
  const selector = `td.sheet-cell[data-row-index=\"${cell.rowIndex}\"][data-col-index=\"${cell.colIndex}\"] select.cell-select`
  const selectElement = document.querySelector(selector)
  if (!(selectElement instanceof HTMLSelectElement)) {
    return
  }
  selectElement.focus()
  selectElement.click()
}

function moveSelectionBy(deltaRow, deltaCol, expand = false) {
  const sheet = currentSheet.value
  if (!sheet) {
    return
  }
  const baseRow = activeSelection.value ? activeSelection.value.startRow : 0
  const baseCol = activeSelection.value ? activeSelection.value.startCol : 0
  const nextRow = clampNumber(baseRow + deltaRow, 0, Math.max(sheet.rowCount - 1, 0))
  const nextCol = clampNumber(baseCol + deltaCol, 0, Math.max(sheet.maxColumnCount - 1, 0))
  selectCoordinate(nextRow, nextCol, expand)
}

function selectColumnDownFromActiveCell() {
  const sheet = currentSheet.value
  const range = activeSelection.value
  if (!sheet || !range) {
    return
  }
  const bottomRow = Math.max(sheet.rowCount - 1, range.endRow)
  if (range.endRow >= bottomRow) {
    return
  }
  selection.value = {
    sheetName: sheet.name,
    anchorStartRow: range.anchorStartRow,
    anchorEndRow: range.anchorEndRow,
    anchorStartCol: range.anchorStartCol,
    anchorEndCol: range.anchorEndCol,
    startRow: range.startRow,
    endRow: bottomRow,
    startCol: range.startCol,
    endCol: range.endCol
  }
}

function selectCoordinate(rowIndex, colIndex, expand = false) {
  const sheet = currentSheet.value
  if (!sheet) {
    return
  }
  const normalizedRowIndex = clampNumber(rowIndex, 0, Math.max(sheet.rowCount - 1, 0))
  const normalizedColIndex = clampNumber(colIndex, 0, Math.max(sheet.maxColumnCount - 1, 0))
  const cell = findCellByCoordinate(sheet, normalizedRowIndex, normalizedColIndex)
  if (!cell) {
    return
  }
  const range = getCellRange(cell)
  if (expand && activeSelection.value) {
    const anchor = activeSelection.value
    selection.value = {
      sheetName: sheet.name,
      anchorStartRow: anchor.anchorStartRow,
      anchorEndRow: anchor.anchorEndRow,
      anchorStartCol: anchor.anchorStartCol,
      anchorEndCol: anchor.anchorEndCol,
      startRow: Math.min(anchor.anchorStartRow, range.startRow),
      endRow: Math.max(anchor.anchorEndRow, range.endRow),
      startCol: Math.min(anchor.anchorStartCol, range.startCol),
      endCol: Math.max(anchor.anchorEndCol, range.endCol)
    }
    return
  }
  selection.value = {
    sheetName: sheet.name,
    anchorStartRow: range.startRow,
    anchorEndRow: range.endRow,
    anchorStartCol: range.startCol,
    anchorEndCol: range.endCol,
    startRow: range.startRow,
    endRow: range.endRow,
    startCol: range.startCol,
    endCol: range.endCol
  }
}

function handleCellFocus(cell) {
  const key = buildCellKey(cell.rowIndex, cell.colIndex)
  if (activeEditCellKey.value === key) {
    return
  }
  pushUndoSnapshot()
  activeEditCellKey.value = key
  formulaInput.value = normalizeValue(cell.value)
}

function handleCellBlur() {
  activeEditCellKey.value = ''
  formulaBarFocused.value = false
}

function clearSelectedCellValues() {
  const sheet = currentSheet.value
  const range = activeSelection.value
  if (!sheet || !range) {
    return
  }
  pushUndoSnapshot()
  const touched = new Set()
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let col = range.startCol; col <= range.endCol; col += 1) {
      const cell = findCellByCoordinate(sheet, row, col)
      if (!cell) {
        continue
      }
      const key = buildCellKey(cell.rowIndex, cell.colIndex)
      if (touched.has(key)) {
        continue
      }
      touched.add(key)
      applyCellInputValue(sheet, cell, '', { skipValidation: true })
    }
  }
  recomputeSheetFormulaDisplays(sheet)
  scheduleDirtyStateRefresh()
}

async function copySelection(cut = false) {
  const sheet = currentSheet.value
  const range = activeSelection.value
  if (!sheet || !range) {
    return
  }
  const richMatrix = buildSelectionSnapshotMatrix(sheet, range)
  const matrix = richMatrix.map((row) => row.map((item) => item.value))
  const text = matrix.map((row) => row.join('\t')).join('\n')
  clipboardPayload.value = {
    mode: cut ? 'cut' : 'copy',
    source: 'excel-editor',
    matrix,
    richMatrix
  }
  await writeClipboardText(text)
      ElMessage.warning('?????????')
    clearSelectedCellValues()
  }
  hideContextMenu()
  ElMessage.success(cut ? '宸插壀鍒囧埌鍓创鏉? : '宸插鍒跺埌鍓创鏉?)
}

async function pasteToSelection() {
  const sheet = currentSheet.value
  const range = activeSelection.value
  if (!sheet || !range) {
    return
  }
  let matrix = Array.isArray(clipboardPayload.value?.matrix) ? clipboardPayload.value.matrix : []
  let richMatrix = Array.isArray(clipboardPayload.value?.richMatrix) ? clipboardPayload.value.richMatrix : []
  if (!matrix.length) {
    const clipboardText = await readClipboardText()
    if (clipboardText) {
      matrix = parseClipboardMatrix(clipboardText)
    }
  }
  if (!matrix.length) {
    ElMessage.warning('鍓创鏉挎病鏈夊彲绮樿创鍐呭')
    return
  }
  pushUndoSnapshot()
  let appliedCount = 0
  let invalidCount = 0
  for (let rowOffset = 0; rowOffset < matrix.length; rowOffset += 1) {
    const rowValues = Array.isArray(matrix[rowOffset]) ? matrix[rowOffset] : []
    const rowSources = Array.isArray(richMatrix[rowOffset]) ? richMatrix[rowOffset] : []
    for (let colOffset = 0; colOffset < rowValues.length; colOffset += 1) {
      const rowIndex = range.startRow + rowOffset
      const colIndex = range.startCol + colOffset
      if (rowIndex >= sheet.rowCount || colIndex >= sheet.maxColumnCount) {
        continue
      }
      let cell = findCellByCoordinate(sheet, rowIndex, colIndex)
      if (!cell) {
        cell = createSyntheticCell(sheet, rowIndex, colIndex, null, false)
        sheet.cells.push(cell)
      }
      const source = normalizeClipboardCellSource(rowSources[colOffset], rowValues[colOffset])
      const applied = applyPastedSourceToCell(sheet, cell, source)
      if (applied) {
        appliedCount += 1
      } else {
        invalidCount += 1
      }
    }
  }
  if (!appliedCount) {
    ElMessage.warning('鐩爣鍖哄煙瓒呭嚭褰撳墠琛ㄦ牸鑼冨洿')
    return
  }
  recomputeSheetFormulaDisplays(sheet)
  rebuildSheetRows(sheet)
  recountSheetCellDirtyCount(sheet)
  scheduleDirtyStateRefresh()
  hideContextMenu()
  if (invalidCount > 0) {
    ElMessage.warning('鏈?' + invalidCount + ' 涓笅鎷夊崟鍏冩牸鍥犻€夐」涓嶅悎娉曡璺宠繃')
  }
}

function buildSelectionMatrix(sheet, range) {
  const matrix = []
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    const rowValues = []
    for (let col = range.startCol; col <= range.endCol; col += 1) {
      const cell = findCellByCoordinate(sheet, row, col)
      rowValues.push(cell ? normalizeValue(cell.value) : '')
    }
    matrix.push(rowValues)
  }
  return matrix
}

function buildSelectionSnapshotMatrix(sheet, range) {
  const matrix = []
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    const rowItems = []
    for (let col = range.startCol; col <= range.endCol; col += 1) {
      const cell = findCellByCoordinate(sheet, row, col)
      rowItems.push({
        value: cell ? normalizeValue(cell.value) : '',
        formula: Boolean(cell?.formula),
        displayValue: cell ? normalizeValue(cell.displayValue) : '',
        style: cell ? cloneCellStyle(cell.style) : {}
      })
    }
    matrix.push(rowItems)
  }
  return matrix
}

function normalizeClipboardCellSource(source, fallbackValue = '') {
  if (!source || typeof source !== 'object') {
    return {
      value: normalizeValue(fallbackValue),
      formula: false,
      style: {}
    }
  }
  return {
    value: normalizeValue(source.value),
    formula: Boolean(source.formula),
    style: cloneCellStyle(source.style)
  }
}

function applyPastedSourceToCell(sheet, cell, source) {
  const nextValue = source?.formula ? normalizeValue(source.value) : normalizeValue(source?.value)
  const result = applyCellInputValue(sheet, cell, nextValue, { skipValidation: false })
  if (!result.applied) {
    return false
  }
  if (source && source.style && typeof source.style === 'object' && Object.keys(source.style).length > 0) {
    applyCellStyleChange(sheet, cell, source.style)
  }
  return true
}

function parseClipboardMatrix(text) {
  return String(text || '')
    .split(/\r?\n/)
    .filter((line) => line.length > 0)
    .map((line) => line.split('\t'))
}

async function writeClipboardText(text) {
  if (!navigator?.clipboard?.writeText) {
    return
  }
  try {
    await navigator.clipboard.writeText(text)
  } catch {
    // Ignore browser clipboard permission failure.
  }
}

async function readClipboardText() {
  if (!navigator?.clipboard?.readText) {
    return ''
  }
  try {
    return await navigator.clipboard.readText()
  } catch {
    return ''
  }
}

function openContextMenu(event, targetCell) {
  const sheet = currentSheet.value
  if (!sheet) {
    return
  }
  const targetRow = clampNumber(targetCell?.rowIndex ?? 0, 0, Math.max(sheet.rowCount - 1, 0))
  const targetCol = clampNumber(targetCell?.colIndex ?? 0, 0, Math.max(sheet.maxColumnCount - 1, 0))
  const resolvedCell = findCellByCoordinate(sheet, targetRow, targetCol)
  if (resolvedCell) {
    handleCellSelection(resolvedCell, {})
  }
  contextMenuState.value = {
    visible: true,
    x: Number(event.clientX || 0),
    y: Number(event.clientY || 0),
    rowIndex: resolvedCell?.rowIndex ?? targetRow,
    colIndex: resolvedCell?.colIndex ?? targetCol
  }
}

function hideContextMenu() {
  if (!contextMenuState.value.visible) {
    return
  }
  contextMenuState.value.visible = false
}

function createSheet() {
  const name = buildUniqueSheetName('Sheet')
  const nextSheet = createEmptySheet(name)
  pushUndoSnapshot()
  sheets.value.push(nextSheet)
  activeSheetName.value = name
  pendingStructures.value.push({
    type: 'sheet',
    action: 'add',
    targetSheetName: name,
    targetIndex: sheets.value.length - 1
  })
  scheduleDirtyStateRefresh()
}

async function renameCurrentSheet() {
  const sheet = currentSheet.value
  if (!sheet) {
    return
  }
  try {
    const { value } = await ElMessageBox.prompt('璇疯緭鍏ユ柊鐨勫伐浣滆〃鍚嶇О', '閲嶅懡鍚嶅伐浣滆〃', {
      confirmButtonText: '纭',
      cancelButtonText: '鍙栨秷',
      inputValue: sheet.name
    })
    const nextName = normalizeValue(value).trim()
    if (!nextName) {
      ElMessage.warning('宸ヤ綔琛ㄥ悕绉颁笉鑳戒负绌?)
      return
    }
    if (sheets.value.some((item) => item.name === nextName && item !== sheet)) {
      ElMessage.warning('宸ヤ綔琛ㄥ悕绉板凡瀛樺湪')
      return
    }
    if (nextName === sheet.name) {
      return
    }
    pushUndoSnapshot()
    const previousName = sheet.name
    sheet.name = nextName
    activeSheetName.value = nextName
    pendingStructures.value.push({
      type: 'sheet',
      action: 'rename',
      sheetName: previousName,
      targetSheetName: nextName
    })
    scheduleDirtyStateRefresh()
  } catch {
    // Ignore cancel action.
  }
}

function copyCurrentSheet() {
  const sheet = currentSheet.value
  if (!sheet) {
    return
  }
  const copyName = buildUniqueSheetName(`${sheet.name}_鍓湰`)
  pushUndoSnapshot()
  const copy = JSON.parse(JSON.stringify(sheet))
  copy.name = copyName
  initializeSheetDirtyMetrics(copy)
  const insertIndex = currentSheetIndex.value + 1
  sheets.value.splice(insertIndex, 0, copy)
  activeSheetName.value = copyName
  pendingStructures.value.push({
    type: 'sheet',
    action: 'copy',
    ElMessage.warning('????????')
    targetSheetName: copyName,
    targetIndex: insertIndex
  })
  scheduleDirtyStateRefresh()
}

function deleteCurrentSheet() {
  if (!canDeleteCurrentSheet.value || !currentSheet.value) {
    return
  }
  const deletingSheet = currentSheet.value
  pushUndoSnapshot()
  const index = currentSheetIndex.value
  sheets.value.splice(index, 1)
  const nextIndex = Math.min(index, sheets.value.length - 1)
  activeSheetName.value = sheets.value[nextIndex]?.name || ''
  pendingStructures.value.push({
    type: 'sheet',
    action: 'delete',
    sheetName: deletingSheet.name
  })
  scheduleDirtyStateRefresh()
}

function moveCurrentSheetLeft() {
  moveCurrentSheetBy(-1)
}

function moveCurrentSheetRight() {
  moveCurrentSheetBy(1)
}

function moveCurrentSheetBy(offset) {
  if (!currentSheet.value) {
    return
  }
  const sourceIndex = currentSheetIndex.value
  const targetIndex = sourceIndex + offset
  if (targetIndex < 0 || targetIndex >= sheets.value.length) {
    return
  }
  pushUndoSnapshot()
  const [sheet] = sheets.value.splice(sourceIndex, 1)
  sheets.value.splice(targetIndex, 0, sheet)
  activeSheetName.value = sheet.name
  pendingStructures.value.push({
    type: 'sheet',
    action: 'reorder',
    sheetName: sheet.name,
    targetIndex
  })
  scheduleDirtyStateRefresh()
}

function buildUniqueSheetName(baseName) {
  const source = normalizeValue(baseName).trim() || 'Sheet'
  if (!sheets.value.some((sheet) => sheet.name === source)) {
    return source
  }
  let index = 1
  while (sheets.value.some((sheet) => sheet.name === `${source}${index}`)) {
    index += 1
  }
  return `${source}${index}`
}

function createEmptySheet(name) {
  const rowCount = DEFAULT_EMPTY_SHEET_ROW_COUNT
  const maxColumnCount = DEFAULT_EMPTY_SHEET_COLUMN_COUNT
  const sheet = {
    name,
    rowCount,
    maxColumnCount,
    rowHeights: Array.from({ length: rowCount }, () => DEFAULT_ROW_HEIGHT),
    columnWidths: Array.from({ length: maxColumnCount }, () => DEFAULT_COLUMN_WIDTH),
    freezePane: normalizeFreezePane(null),
    filterConfig: normalizeFilterConfig(null),
    validations: [],
    images: [],
    cells: [],
    rows: [],
    originalMergeKeys: [],
    originalRowHeights: Array.from({ length: rowCount }, () => DEFAULT_ROW_HEIGHT),
    originalColumnWidths: Array.from({ length: maxColumnCount }, () => DEFAULT_COLUMN_WIDTH)
  }
  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    for (let colIndex = 0; colIndex < maxColumnCount; colIndex += 1) {
      sheet.cells.push({
        rowIndex,
        colIndex,
        rowSpan: 1,
        colSpan: 1,
        widthPx: DEFAULT_COLUMN_WIDTH,
        heightPx: DEFAULT_ROW_HEIGHT,
        value: '',
        originalValue: '',
        displayValue: '',
        formula: false,
    ElMessage.warning('????????')
        forceDirty: false,
        validationOptions: [],
        validationAllowBlank: true,
        style: {},
        originalStyle: {}
      })
    }
  }
  rebuildSheetRows(sheet)
  initializeSheetDirtyMetrics(sheet)
  return sheet
}

function insertRowAbove() {
  const range = activeSelection.value
  if (!range) {
    return
  }
  insertRowsAt(range.startRow, range.endRow - range.startRow + 1)
  hideContextMenu()
}

function insertRowBelow() {
  const range = activeSelection.value
  if (!range) {
    return
  }
  insertRowsAt(range.endRow + 1, range.endRow - range.startRow + 1)
  hideContextMenu()
}

function deleteSelectedRows() {
  const range = activeSelection.value
  if (!range) {
    return
  }
  deleteRowsAt(range.startRow, range.endRow - range.startRow + 1)
  hideContextMenu()
}

function insertColumnLeft() {
  const range = activeSelection.value
  if (!range) {
    return
  }
  insertColumnsAt(range.startCol, range.endCol - range.startCol + 1)
  hideContextMenu()
}

function insertColumnRight() {
  const range = activeSelection.value
  if (!range) {
    return
  }
  insertColumnsAt(range.endCol + 1, range.endCol - range.startCol + 1)
  hideContextMenu()
}

function deleteSelectedColumns() {
  const range = activeSelection.value
  if (!range) {
    return
  }
  deleteColumnsAt(range.startCol, range.endCol - range.startCol + 1)
  hideContextMenu()
}

function insertRowsAt(startIndex, count) {
  const sheet = currentSheet.value
  if (!sheet || count <= 0) {
    return
  }
  const insertIndex = clampNumber(startIndex, 0, Math.max(sheet.rowCount, 0))
  pushUndoSnapshot()
  for (const cell of sheet.cells) {
    const cellEnd = cell.rowIndex + cell.rowSpan - 1
    if (cell.rowIndex >= insertIndex) {
      cell.rowIndex += count
      continue
    }
    if (cellEnd >= insertIndex) {
      cell.rowSpan += count
    }
  }
  for (let offset = 0; offset < count; offset += 1) {
    sheet.rowHeights.splice(insertIndex, 0, DEFAULT_ROW_HEIGHT)
    sheet.originalRowHeights.splice(insertIndex, 0, DEFAULT_ROW_HEIGHT)
  }
  sheet.rowCount += count
  pendingStructures.value.push({
    type: 'row',
    action: 'insert',
    sheetName: sheet.name,
    index: insertIndex,
    count
  })
  ensureSheetDenseGrid(sheet)
  recountSheetCellDirtyCount(sheet)
  recountSheetMergeDirtyCount(sheet)
  rebuildSheetDimensionDirtyIndexes(sheet)
  resetSheetFilterOnStructureChange(sheet)
  scheduleDirtyStateRefresh()
}

function deleteRowsAt(startIndex, count) {
  const sheet = currentSheet.value
  if (!sheet || count <= 0 || sheet.rowCount <= count) {
    ElMessage.warning('鑷冲皯淇濈暀涓€琛屾暟鎹?)
    return
  }
  const deleteStart = clampNumber(startIndex, 0, Math.max(sheet.rowCount - 1, 0))
  const deleteCount = Math.min(count, sheet.rowCount - 1 - deleteStart + 1)
  const deleteEnd = deleteStart + deleteCount - 1
  pushUndoSnapshot()
  const nextCells = []
  for (const cell of sheet.cells) {
    const cellStart = cell.rowIndex
    const cellEnd = cell.rowIndex + cell.rowSpan - 1
    if (cellEnd < deleteStart) {
      nextCells.push(cell)
      continue
    }
    if (cellStart > deleteEnd) {
      cell.rowIndex -= deleteCount
      nextCells.push(cell)
      continue
    }
    if (cellStart < deleteStart && cellEnd > deleteEnd) {
      cell.rowSpan -= deleteCount
      nextCells.push(cell)
      continue
    }
    if (cellStart < deleteStart && cellEnd >= deleteStart) {
      cell.rowSpan = deleteStart - cellStart
      if (cell.rowSpan > 0) {
        nextCells.push(cell)
      }
      continue
    }
    if (cellStart <= deleteEnd && cellEnd > deleteEnd) {
      cell.rowIndex = deleteStart
      cell.rowSpan = cellEnd - deleteEnd
      if (cell.rowSpan > 0) {
        nextCells.push(cell)
      }
    }
  }
  sheet.cells = dedupeCells(nextCells)
  sheet.rowHeights.splice(deleteStart, deleteCount)
  sheet.originalRowHeights.splice(deleteStart, deleteCount)
  sheet.rowCount -= deleteCount
  pendingStructures.value.push({
    type: 'row',
    action: 'delete',
    sheetName: sheet.name,
    index: deleteStart,
    count: deleteCount
  })
  ensureSheetDenseGrid(sheet)
  recountSheetCellDirtyCount(sheet)
  recountSheetMergeDirtyCount(sheet)
  rebuildSheetDimensionDirtyIndexes(sheet)
  resetSheetFilterOnStructureChange(sheet)
  scheduleDirtyStateRefresh()
}

function insertColumnsAt(startIndex, count) {
  const sheet = currentSheet.value
  if (!sheet || count <= 0) {
    return
  }
  const insertIndex = clampNumber(startIndex, 0, Math.max(sheet.maxColumnCount, 0))
  pushUndoSnapshot()
  for (const cell of sheet.cells) {
    const cellEnd = cell.colIndex + cell.colSpan - 1
    if (cell.colIndex >= insertIndex) {
      cell.colIndex += count
      continue
    }
    if (cellEnd >= insertIndex) {
      cell.colSpan += count
    }
  }
  for (let offset = 0; offset < count; offset += 1) {
    sheet.columnWidths.splice(insertIndex, 0, DEFAULT_COLUMN_WIDTH)
    sheet.originalColumnWidths.splice(insertIndex, 0, DEFAULT_COLUMN_WIDTH)
  }
  sheet.maxColumnCount += count
  pendingStructures.value.push({
    type: 'column',
    action: 'insert',
    sheetName: sheet.name,
    index: insertIndex,
    count
  })
  ensureSheetDenseGrid(sheet)
  recountSheetCellDirtyCount(sheet)
  recountSheetMergeDirtyCount(sheet)
  rebuildSheetDimensionDirtyIndexes(sheet)
  resetSheetFilterOnStructureChange(sheet)
  scheduleDirtyStateRefresh()
}

function deleteColumnsAt(startIndex, count) {
  const sheet = currentSheet.value
  if (!sheet || count <= 0 || sheet.maxColumnCount <= count) {
    ElMessage.warning('鑷冲皯淇濈暀涓€鍒楁暟鎹?)
    return
  }
  const deleteStart = clampNumber(startIndex, 0, Math.max(sheet.maxColumnCount - 1, 0))
  const deleteCount = Math.min(count, sheet.maxColumnCount - 1 - deleteStart + 1)
  const deleteEnd = deleteStart + deleteCount - 1
  pushUndoSnapshot()
  const nextCells = []
  for (const cell of sheet.cells) {
    const cellStart = cell.colIndex
    const cellEnd = cell.colIndex + cell.colSpan - 1
    if (cellEnd < deleteStart) {
      nextCells.push(cell)
      continue
    }
    if (cellStart > deleteEnd) {
      cell.colIndex -= deleteCount
      nextCells.push(cell)
      continue
    }
    if (cellStart < deleteStart && cellEnd > deleteEnd) {
      cell.colSpan -= deleteCount
      nextCells.push(cell)
      continue
    }
    if (cellStart < deleteStart && cellEnd >= deleteStart) {
      cell.colSpan = deleteStart - cellStart
      if (cell.colSpan > 0) {
        nextCells.push(cell)
      }
      continue
    }
    if (cellStart <= deleteEnd && cellEnd > deleteEnd) {
      cell.colIndex = deleteStart
      cell.colSpan = cellEnd - deleteEnd
      if (cell.colSpan > 0) {
        nextCells.push(cell)
      }
    }
  }
  sheet.cells = dedupeCells(nextCells)
  sheet.columnWidths.splice(deleteStart, deleteCount)
  sheet.originalColumnWidths.splice(deleteStart, deleteCount)
  sheet.maxColumnCount -= deleteCount
  pendingStructures.value.push({
    type: 'column',
    action: 'delete',
    sheetName: sheet.name,
    index: deleteStart,
    count: deleteCount
  })
  ensureSheetDenseGrid(sheet)
  recountSheetCellDirtyCount(sheet)
  recountSheetMergeDirtyCount(sheet)
  rebuildSheetDimensionDirtyIndexes(sheet)
  resetSheetFilterOnStructureChange(sheet)
  scheduleDirtyStateRefresh()
}

function toggleBold() {
  toggleStyleToken('fontWeight', '700')
}

function toggleItalic() {
  toggleStyleToken('fontStyle', 'italic')
}

function toggleUnderline() {
  toggleStyleToken('textDecoration', 'underline')
}

function toggleStyleToken(key, token) {
  const sheet = currentSheet.value
  const range = activeSelection.value
  if (!sheet || !range) {
    return
  }
  const cells = getSelectionTopLeftCells(sheet, range)
  if (!cells.length) {
    return
  }
  const allEnabled = cells.every((cell) => normalizeEditableStyleValue(key, cell?.style?.[key]) === token)
  applyStyleToSelection({ [key]: allEnabled ? '' : token })
}

function applyFontSize(value) {
  fontSizePicker.value = normalizeEditableStyleValue('fontSize', value) || '12pt'
  applyStyleToSelection({ fontSize: value })
}

function applyHorizontalAlign(align) {
  applyStyleToSelection({ textAlign: align })
}

function applyVerticalAlign(align) {
  applyStyleToSelection({ verticalAlign: align })
}

function handleTextColorModelValueChange(color) {
  applyTextColorChange(color)
}

function handleTextColorChange(color) {
  applyTextColorChange(color)
}

function clearTextColor() {
  textColorPicker.value = '#111827'
  applyStyleToSelection({ color: '' })
}

function handleFillColorModelValueChange(color) {
  applyFillColorChange(color)
}

function handleFillColorChange(color) {
  applyFillColorChange(color)
}

function clearFillColor() {
  fillColorPicker.value = '#FFFFFF'
  applyStyleToSelection({ backgroundColor: '' })
}

function clearFormatPainterState() {
  formatPainterActive.value = false
  formatPainterStyle.value = null
}

function toggleFormatPainter() {
  if (formatPainterActive.value) {
    clearFormatPainterState()
    return
  }
  const sourceCell = activePrimaryCell.value
  if (!sourceCell) {
    ElMessage.warning('璇峰厛閫夋嫨涓€涓甫鏍煎紡鐨勫崟鍏冩牸')
    return
  }
  formatPainterStyle.value = cloneCellStyle(sourceCell.style)
  formatPainterActive.value = true
  ElMessage.success('鏍煎紡鍒峰凡寮€鍚紝璇风偣鍑荤洰鏍囧崟鍏冩牸')
}

    ElMessage.info('???????????????')
  if (!formatPainterActive.value || !sheet || !range) {
    return
  }
  const stylePatch = cloneCellStyle(formatPainterStyle.value)
  clearFormatPainterState()
  if (!stylePatch || typeof stylePatch !== 'object' || !Object.keys(stylePatch).length) {
    return
  }
  applyStylePatchToRange(sheet, range, stylePatch)
}

function applyBorderToSelection(mode = 'all') {
  const sheet = currentSheet.value
  const range = activeSelection.value
  if (!sheet || !range) {
    return
  }
  const cells = getSelectionTopLeftCells(sheet, range)
  if (!cells.length) {
    return
  }
  let changed = 0
  let pushed = false
  for (const cell of cells) {
    const cellRange = getCellRange(cell)
    const patch = {}
    if (mode === 'clear') {
      patch.borderTop = ''
      patch.borderRight = ''
      patch.borderBottom = ''
      patch.borderLeft = ''
    } else if (mode === 'all') {
      patch.borderTop = BORDER_LINE_STYLE
      patch.borderRight = BORDER_LINE_STYLE
      patch.borderBottom = BORDER_LINE_STYLE
      patch.borderLeft = BORDER_LINE_STYLE
    } else if (mode === 'outer') {
      if (cellRange.startRow === range.startRow) {
        patch.borderTop = BORDER_LINE_STYLE
      }
      if (cellRange.endRow === range.endRow) {
        patch.borderBottom = BORDER_LINE_STYLE
      }
      if (cellRange.startCol === range.startCol) {
        patch.borderLeft = BORDER_LINE_STYLE
      }
      if (cellRange.endCol === range.endCol) {
        patch.borderRight = BORDER_LINE_STYLE
      }
    }
    if (!Object.keys(patch).length || !willCellStyleChange(cell, patch)) {
      continue
    }
    if (!pushed) {
      pushUndoSnapshot()
      pushed = true
    }
    applyCellStyleChange(sheet, cell, patch)
    changed += 1
  }
  if (!changed) {
    return
  }
  recountSheetCellDirtyCount(sheet)
  scheduleDirtyStateRefresh()
}

function formatSelectionAsPercent() {
  applyNumericTransformToSelection((valueInfo) => ({
    value: `${(valueInfo.value * 100).toFixed(Math.max(valueInfo.decimals, 2))}%`,
    stylePatch: { numberFormat: 'percent' }
  }))
}

function formatSelectionAsDate() {
  const sheet = currentSheet.value
  const range = activeSelection.value
  if (!sheet || !range) {
    return
  }
  const cells = getSelectionTopLeftCells(sheet, range)
  if (!cells.length) {
    return
  }
  let appliedCount = 0
  let skippedCount = 0
  let pushed = false
  for (const cell of cells) {
    if (cell.formula) {
      skippedCount += 1
      continue
    }
    ElMessage.warning('????????????????')
    if (!nextDate) {
      skippedCount += 1
      continue
    }
    const nextText = formatDateValue(nextDate)
    const valueChanged = normalizeValue(cell.value) !== nextText
    const stylePatch = { numberFormat: 'date' }
    const styleChanged = willCellStyleChange(cell, stylePatch)
    if (!valueChanged && !styleChanged) {
      continue
    }
    if (!pushed) {
      pushUndoSnapshot()
      pushed = true
    }
    if (valueChanged) {
      const result = applyCellInputValue(sheet, cell, nextText, { skipValidation: true })
      if (!result.applied) {
        skippedCount += 1
        continue
      }
    }
    if (styleChanged) {
      applyCellStyleChange(sheet, cell, stylePatch)
    }
    appliedCount += 1
  }
  if (!appliedCount) {
    ElMessage.info('褰撳墠閫夊尯娌℃湁鍙浆鎹负鏃ユ湡鐨勫崟鍏冩牸')
    return
  }
  recomputeSheetFormulaDisplays(sheet)
  recountSheetCellDirtyCount(sheet)
  scheduleDirtyStateRefresh()
  if (skippedCount > 0) {
    ElMessage.success(`宸插鐞?${appliedCount} 涓崟鍏冩牸锛岃烦杩?${skippedCount} 涓猔)
  }
}

function increaseSelectionDecimals() {
  applyNumericTransformToSelection((valueInfo) => {
    const decimals = Math.min(valueInfo.decimals + 1, MAX_DECIMAL_SCALE)
    return {
      value: formatNumericByDecimals(valueInfo, decimals),
      stylePatch: { numberFormat: valueInfo.isPercent ? 'percent' : 'number' }
    }
  })
}

function decreaseSelectionDecimals() {
  applyNumericTransformToSelection((valueInfo) => {
    const decimals = Math.max(valueInfo.decimals - 1, 0)
    return {
      value: formatNumericByDecimals(valueInfo, decimals),
      stylePatch: { numberFormat: valueInfo.isPercent ? 'percent' : 'number' }
    }
  })
}

function applyNumericTransformToSelection(transformer) {
  const sheet = currentSheet.value
  const range = activeSelection.value
  if (!sheet || !range || typeof transformer !== 'function') {
    return
  }
  const cells = getSelectionTopLeftCells(sheet, range)
  if (!cells.length) {
    return
  }
  let appliedCount = 0
  let skippedCount = 0
  let pushed = false
  for (const cell of cells) {
    if (cell.formula) {
      skippedCount += 1
      continue
    }
    const valueInfo = parseNumericCellValue(cell.value)
    if (!valueInfo) {
      skippedCount += 1
      continue
    }
    const transformed = transformer(valueInfo, cell)
    if (!transformed || typeof transformed.value === 'undefined') {
      skippedCount += 1
      continue
    }
    const stylePatch = transformed.stylePatch && typeof transformed.stylePatch === 'object'
      ? transformed.stylePatch
      : {}
    const nextValue = normalizeValue(transformed.value)
    const valueChanged = normalizeValue(cell.value) !== nextValue
    const styleChanged = willCellStyleChange(cell, stylePatch)
    if (!valueChanged && !styleChanged) {
      continue
    }
    if (!pushed) {
      pushUndoSnapshot()
      pushed = true
    }
    if (valueChanged) {
      const result = applyCellInputValue(sheet, cell, nextValue, { skipValidation: true })
      if (!result.applied) {
        skippedCount += 1
        continue
      }
    }
    if (styleChanged) {
      applyCellStyleChange(sheet, cell, stylePatch)
    }
    appliedCount += 1
  }
  if (!appliedCount) {
    ElMessage.info('褰撳墠閫夊尯娌℃湁鍙鐞嗙殑鏁板瓧鍗曞厓鏍?)
    return
  }
  recomputeSheetFormulaDisplays(sheet)
  recountSheetCellDirtyCount(sheet)
  scheduleDirtyStateRefresh()
  if (skippedCount > 0) {
    ElMessage.success(`宸插鐞?${appliedCount} 涓崟鍏冩牸锛岃烦杩?${skippedCount} 涓猔)
  }
}

function parseNumericCellValue(value) {
  const text = normalizeValue(value).trim()
  if (!text) {
    return null
  }
  const isPercent = text.endsWith('%')
  const rawNumber = isPercent ? text.slice(0, -1).trim() : text
  const normalized = rawNumber.replace(/,/g, '')
  const parsed = Number(normalized)
  if (!Number.isFinite(parsed)) {
    return null
  }
  const decimals = detectDecimalPlaces(rawNumber)
  return {
    source: text,
    value: isPercent ? parsed / 100 : parsed,
    decimals,
    isPercent
  }
}

function detectDecimalPlaces(text) {
  const source = normalizeValue(text).trim()
  const dotIndex = source.lastIndexOf('.')
  if (dotIndex < 0) {
    return 0
  }
  const decimalPart = source.slice(dotIndex + 1).replace(/[^0-9]/g, '')
  return Math.min(decimalPart.length, MAX_DECIMAL_SCALE)
}

function formatNumericByDecimals(valueInfo, decimals) {
  const safeDecimals = Math.min(Math.max(Number(decimals ?? 0), 0), MAX_DECIMAL_SCALE)
  if (valueInfo.isPercent) {
    return `${(valueInfo.value * 100).toFixed(safeDecimals)}%`
  }
  return Number(valueInfo.value).toFixed(safeDecimals)
}

function parseDateCellValue(value) {
  const text = normalizeValue(value).trim()
  if (!text) {
    return null
  }
  if (/^\d{1,7}$/.test(text)) {
    const serial = Number(text)
    if (Number.isFinite(serial) && serial > 0) {
      return excelSerialToDate(serial)
    }
  }
  const normalized = text.replace(/[./]/g, '-')
  const date = new Date(normalized)
  if (Number.isNaN(date.getTime())) {
    return null
  }
  return date
}

function excelSerialToDate(serial) {
  const base = Date.UTC(1899, 11, 30)
  const milliseconds = Math.round(Number(serial) * 24 * 60 * 60 * 1000)
  return new Date(base + milliseconds)
}

function formatDateValue(date) {
  const year = date.getFullYear()
  const month = String(date.getMonth() + 1).padStart(2, '0')
  const day = String(date.getDate()).padStart(2, '0')
  return `${year}-${month}-${day}`
}

function freezeFirstRow() {
  applyFreezePane(1, Math.max(Number(currentSheet.value?.freezePane?.xSplit ?? 0), 0))
}

function freezeFirstColumn() {
  applyFreezePane(Math.max(Number(currentSheet.value?.freezePane?.ySplit ?? 0), 0), 1)
}

function freezeBySelection() {
  const range = activeSelection.value
  if (!range) {
    ElMessage.warning('璇峰厛閫夋嫨涓€涓崟鍏冩牸浣滀负鍐荤粨鍒嗗壊鐐?)
    return
  }
  applyFreezePane(range.startRow, range.startCol)
}

function openFreezeDialog() {
  const sheet = currentSheet.value
  if (!sheet) {
    return
  }
  freezeForm.value = {
    rows: Math.max(Number(sheet.freezePane?.ySplit ?? 0), 0),
    cols: Math.max(Number(sheet.freezePane?.xSplit ?? 0), 0)
  }
  freezeDialogVisible.value = true
}

function applyCustomFreezePane() {
  const rows = Math.max(Number(freezeForm.value.rows ?? 0), 0)
  const cols = Math.max(Number(freezeForm.value.cols ?? 0), 0)
  applyFreezePane(rows, cols)
  freezeDialogVisible.value = false
}

function clearFreezePane() {
  applyFreezePane(0, 0)
}

function applyFreezePane(rows, cols) {
  const sheet = currentSheet.value
  if (!sheet) {
    return
  }
  ensureSheetBounds(sheet, sheet.rowCount, sheet.maxColumnCount)
  const nextRows = clampNumber(rows, 0, Math.max(sheet.rowCount - 1, 0))
  const nextCols = clampNumber(cols, 0, Math.max(sheet.maxColumnCount - 1, 0))
  const currentRows = Math.max(Number(sheet.freezePane?.ySplit ?? 0), 0)
  const currentCols = Math.max(Number(sheet.freezePane?.xSplit ?? 0), 0)
  if (currentRows === nextRows && currentCols === nextCols) {
    return
  }
  sheet.freezePane = normalizeFreezePane({
    xSplit: nextCols,
    ySplit: nextRows,
    leftColumn: nextCols,
    topRow: nextRows
  })
  scheduleDirtyStateRefresh()
}

function applyTextColorChange(color) {
  const resolved = resolveColorPickerValue(color, '#111827')
  textColorPicker.value = resolved.pickerValue
  if (!resolved.valid) {
    return
  }
    ElMessage.warning('???????????')
}

function applyFillColorChange(color) {
  const resolved = resolveColorPickerValue(color, '#FFFFFF')
  fillColorPicker.value = resolved.pickerValue
  if (!resolved.valid) {
    return
  }
  applyStyleToSelection({ backgroundColor: resolved.styleValue })
}

function resolveColorPickerValue(color, fallbackColor) {
  const raw = normalizeValue(color).trim()
  if (!raw) {
    return {
      valid: true,
      pickerValue: fallbackColor,
      styleValue: ''
    }
  }
  const normalized = normalizeColorStyleValue(raw)
  if (!normalized) {
    return {
      valid: false,
      pickerValue: fallbackColor,
      styleValue: ''
    }
  }
    ElMessage.warning('???????????')
    valid: true,
    pickerValue: normalized,
    styleValue: normalized
  }
}

function applyStyleToSelection(stylePatch) {
  const sheet = currentSheet.value
  const range = activeSelection.value
  if (!sheet || !range || !stylePatch || typeof stylePatch !== 'object') {
    return
  }
  applyStylePatchToRange(sheet, range, stylePatch)
}

function applyStylePatchToRange(sheet, range, stylePatch) {
  if (!sheet || !range || !stylePatch || typeof stylePatch !== 'object') {
    return
  }
  const cells = getSelectionTopLeftCells(sheet, range)
  if (!cells.length) {
    return
  }
  const changedCells = cells.filter((cell) => willCellStyleChange(cell, stylePatch))
  if (!changedCells.length) {
    return
  }
  pushUndoSnapshot()
  for (const cell of changedCells) {
    applyCellStyleChange(sheet, cell, stylePatch)
  }
  recountSheetCellDirtyCount(sheet)
  scheduleDirtyStateRefresh()
}

function willCellStyleChange(cell, stylePatch) {
  if (!cell || !stylePatch || typeof stylePatch !== 'object') {
    return false
  }
  const sourceStyle = cell.style && typeof cell.style === 'object' ? cell.style : {}
  for (const [styleKey, styleValue] of Object.entries(stylePatch)) {
    if (!EDITABLE_STYLE_KEYS.includes(styleKey)) {
      continue
    }
    const current = normalizeEditableStyleValue(styleKey, sourceStyle[styleKey])
    const next = normalizeEditableStyleValue(styleKey, styleValue)
    if (current !== next) {
      return true
    }
  }
  return false
}

function getSelectionTopLeftCells(sheet, range) {
  const list = []
  const dedupe = new Set()
  for (let rowIndex = range.startRow; rowIndex <= range.endRow; rowIndex += 1) {
    for (let colIndex = range.startCol; colIndex <= range.endCol; colIndex += 1) {
      const cell = findCellByCoordinate(sheet, rowIndex, colIndex)
      if (!cell) {
        continue
      }
      const key = buildCellKey(cell.rowIndex, cell.colIndex)
      if (dedupe.has(key)) {
        continue
      }
      dedupe.add(key)
      list.push(cell)
    }
  }
  return list
}

function sortSelection(direction = 'asc') {
  const range = activeSelection.value
  if (!range) {
    return
  }
  sortRangeByRules(range, [{ colIndex: range.startCol, direction }], {
    hasHeader: sortHasHeader.value,
    showMessage: true
  })
}

function openSortDialog() {
  const range = activeSelection.value
  if (!range) {
    ElMessage.warning('璇峰厛閫夋嫨闇€瑕佹帓搴忕殑鍖哄煙')
    return
  }
  sortDialogRange.value = { ...range }
  sortDialogHasHeader.value = Boolean(sortHasHeader.value)
  sortRules.value = [{ colIndex: range.startCol, direction: 'asc' }]
  sortDialogVisible.value = true
}

function addSortRule() {
  const range = sortDialogRange.value
  if (!range) {
    return
  }
  const used = new Set(sortRules.value.map((rule) => Number(rule.colIndex)))
  let candidate = range.startCol
  while (candidate <= range.endCol && used.has(candidate)) {
    candidate += 1
  }
  if (candidate > range.endCol) {
    candidate = range.startCol
  }
  sortRules.value.push({
    colIndex: candidate,
    direction: 'asc'
  })
}

function removeSortRule(ruleIndex) {
  if (sortRules.value.length <= 1) {
    return
  }
  sortRules.value.splice(ruleIndex, 1)
}

function applySortRules() {
  const range = sortDialogRange.value
  if (!range) {
    return
  }
  const success = sortRangeByRules(range, sortRules.value, {
    hasHeader: sortDialogHasHeader.value,
    showMessage: true
  })
  if (!success) {
    return
  ElMessage.success('???????')
  sortHasHeader.value = sortDialogHasHeader.value
  sortDialogVisible.value = false
}

function normalizeSortRules(rules, range) {
  const list = Array.isArray(rules) ? rules : []
  const normalized = []
  const used = new Set()
  for (const rule of list) {
    const colIndex = Number(rule?.colIndex)
    if (!Number.isFinite(colIndex) || colIndex < range.startCol || colIndex > range.endCol || used.has(colIndex)) {
  ElMessage.success('?????')
    }
    used.add(colIndex)
    normalized.push({
      colIndex,
      direction: normalizeSortDirection(rule?.direction)
    })
  }
  return normalized
}

function normalizeSortDirection(direction) {
  return normalizeValue(direction).toLowerCase() === 'desc' ? 'desc' : 'asc'
}

function sortRangeByRules(range, rules, options = {}) {
  const sheet = currentSheet.value
  if (!sheet || !range) {
    return false
  }
  if (sheet.cells.some((cell) => isMergedCell(cell) && rangesIntersect(range, getCellRange(cell)))) {
    ElMessage.warning('鍖呭惈鍚堝苟鍗曞厓鏍肩殑鍖哄煙鏆備笉鏀寔鎺掑簭')
    return false
  }
  const hasHeader = options?.hasHeader !== false
    ElMessage.warning('??????')
  if (dataStartRow > range.endRow) {
    ElMessage.warning('璇疯嚦灏戦€夋嫨涓€琛屽彲鎺掑簭鏁版嵁')
    return false
  }
  const normalizedRules = normalizeSortRules(rules, range)
  if (!normalizedRules.length) {
    ElMessage.warning('璇疯嚦灏戦厤缃竴鍒楁帓搴忚鍒?)
    return false
  }
  const columnIndexes = []
  for (let colIndex = range.startCol; colIndex <= range.endCol; colIndex += 1) {
    columnIndexes.push(colIndex)
  }

  const sortableRows = []
  for (let rowIndex = dataStartRow; rowIndex <= range.endRow; rowIndex += 1) {
    const snapshots = []
    ElMessage.warning('??????')
    for (const colIndex of columnIndexes) {
      let cell = findCellByCoordinate(sheet, rowIndex, colIndex)
      if (!cell) {
        cell = createSyntheticCell(sheet, rowIndex, colIndex, null, false)
        sheet.cells.push(cell)
      }
      sortValues[colIndex] = normalizeValue(cell.value)
      snapshots.push(snapshotSortableCell(cell))
    }
    sortableRows.push({
      originalRowIndex: rowIndex,
      sortValues,
      snapshots
    })
  }

  if (!sortableRows.length) {
    ElMessage.warning('褰撳墠閫夊尯娌℃湁鍙帓搴忔暟鎹?)
    return false
  }

  pushUndoSnapshot()
  sortableRows.sort((left, right) => {
    for (const rule of normalizedRules) {
      const compareResult = compareSortValues(
        left.sortValues[rule.colIndex],
        right.sortValues[rule.colIndex],
        rule.direction
      )
      if (compareResult !== 0) {
        return compareResult
      }
    }
    return left.originalRowIndex - right.originalRowIndex
  })

  for (let rowOffset = 0; rowOffset < sortableRows.length; rowOffset += 1) {
    const targetRowIndex = dataStartRow + rowOffset
    const rowSnapshot = sortableRows[rowOffset]
    for (let colOffset = 0; colOffset < columnIndexes.length; colOffset += 1) {
      const targetColIndex = columnIndexes[colOffset]
      let targetCell = findCellByCoordinate(sheet, targetRowIndex, targetColIndex)
      if (!targetCell) {
        targetCell = createSyntheticCell(sheet, targetRowIndex, targetColIndex, null, false)
        sheet.cells.push(targetCell)
      }
      applySortableSnapshotToCell(sheet, targetCell, rowSnapshot.snapshots[colOffset])
    }
  }

  recomputeSheetFormulaDisplays(sheet)
  rebuildSheetRows(sheet)
  recountSheetCellDirtyCount(sheet)
  scheduleDirtyStateRefresh()
  if (options?.showMessage) {
    ElMessage.success('鎺掑簭瀹屾垚')
  }
  return true
}

function snapshotSortableCell(cell) {
  return {
    value: normalizeValue(cell?.value),
    displayValue: normalizeValue(cell?.displayValue),
    formula: Boolean(cell?.formula),
    style: cloneCellStyle(cell?.style)
  }
}


function applySortableSnapshotToCell(sheet, cell, snapshot) {
  const next = snapshot || { value: '', displayValue: '', formula: false, style: {} }
  cell.formula = Boolean(next.formula)
  cell.displayValue = normalizeValue(next.displayValue || next.value)
  cell.style = cloneCellStyle(next.style)
  applyCellDirtyChange(sheet, cell, normalizeValue(next.value))
}

function compareSortValues(leftValue, rightValue, direction) {
  const leftText = normalizeValue(leftValue).trim()
  const rightText = normalizeValue(rightValue).trim()
  const leftEmpty = leftText === ''
  const rightEmpty = rightText === ''

  if (leftEmpty && rightEmpty) {
    return 0
  }
  if (leftEmpty) {
    return 1
  }
  if (rightEmpty) {
    return -1
  }

  const leftNumber = Number(leftText)
  const rightNumber = Number(rightText)
  const leftIsNumber = !Number.isNaN(leftNumber)
  const rightIsNumber = !Number.isNaN(rightNumber)

  let compareResult = 0
  if (leftIsNumber && rightIsNumber) {
    compareResult = leftNumber - rightNumber
  } else {
    compareResult = leftText.localeCompare(rightText, 'zh-CN', { numeric: true, sensitivity: 'base' })
  }
  return direction === 'desc' ? -compareResult : compareResult
}

function isColumnInsideFilterRange(colIndex, filterConfig) {
  const normalized = normalizeFilterConfig(filterConfig)
  return colIndex >= normalized.startCol && colIndex <= normalized.endCol
}

function toggleAutoFilter() {
  const sheet = currentSheet.value
  const range = activeSelection.value
  if (!sheet || !range) {
    ElMessage.warning('璇峰厛閫夋嫨闇€瑕佺瓫閫夌殑鍖哄煙')
    return
  }
  const hasHeader = sortHasHeader.value !== false
  const dataStartRow = hasHeader ? range.startRow + 1 : range.startRow
  if (dataStartRow > range.endRow) {
    ElMessage.warning(hasHeader ? '绛涢€夊尯鍩熻嚦灏戦渶瑕佷竴琛岃〃澶村拰涓€琛屾暟鎹? : '鏃犺〃澶存ā寮忎笅鑷冲皯闇€瑕佷竴琛屽彲绛涢€夋暟鎹?)
    return
  }
  const currentFilter = normalizeFilterConfig(sheet.filterConfig)
  const sameRange = currentFilter.enabled
    && currentFilter.hasHeader === hasHeader
    && currentFilter.headerRow === range.startRow
    && currentFilter.startCol === range.startCol
    && currentFilter.endCol === range.endCol
    && currentFilter.dataStartRow === dataStartRow
    && currentFilter.dataEndRow === range.endRow
  if (sameRange) {
    clearAutoFilter()
    return
  }
  sheet.filterConfig = normalizeFilterConfig({
    enabled: true,
    hasHeader,
    headerRow: range.startRow,
    startCol: range.startCol,
    endCol: range.endCol,
    dataStartRow,
    dataEndRow: range.endRow,
    criteria: {}
  })
  filterDialogColIndex.value = range.startCol
  filterKeyword.value = ''
  filterSelectedValues.value = []
  ElMessage.success('宸插紑鍚嚜鍔ㄧ瓫閫?)
}

function clearAutoFilter() {
  const sheet = currentSheet.value
  if (!sheet) {
    return
  }
  sheet.filterConfig = normalizeFilterConfig(null)
  columnFilterDialogVisible.value = false
  filterKeyword.value = ''
  filterSelectedValues.value = []
  ElMessage.success('宸叉竻闄ょ瓫閫?)
}

function isFilterHeaderCell(cell) {
  const sheet = currentSheet.value
  if (!sheet || !cell) {
    return false
  }
  const filterConfig = normalizeFilterConfig(sheet.filterConfig)
  if (!filterConfig.enabled) {
    return false
  }
  if (cell.rowIndex !== filterConfig.headerRow) {
    return false
  }
  return isColumnInsideFilterRange(cell.colIndex, filterConfig)
}

function openColumnFilterDialogBySelection() {
  const sheet = currentSheet.value
  if (!sheet) {
    return
  }
  const filterConfig = normalizeFilterConfig(sheet.filterConfig)
  if (!filterConfig.enabled) {
    ElMessage.warning('璇峰厛寮€鍚瓫閫?)
    return
  }
  const range = activeSelection.value
  if (range && isColumnInsideFilterRange(range.startCol, filterConfig)) {
    openColumnFilterDialog(range.startCol)
    return
  }
  openColumnFilterDialog(filterConfig.startCol)
}

function openColumnFilterDialog(colIndex) {
  const sheet = currentSheet.value
  if (!sheet) {
    return
  }
  const filterConfig = normalizeFilterConfig(sheet.filterConfig)
  if (!filterConfig.enabled) {
    ElMessage.warning('璇峰厛寮€鍚瓫閫?)
    return
  }
  if (!isColumnInsideFilterRange(colIndex, filterConfig)) {
    ElMessage.warning('璇烽€夋嫨绛涢€夊尯鍩熷唴鐨勫垪')
    return
  }
  filterDialogColIndex.value = colIndex
  filterKeyword.value = ''
  resetFilterCandidateValues()
  columnFilterDialogVisible.value = true
}

function resetFilterCandidateValues() {
  const sheet = currentSheet.value
  if (!sheet || filterDialogColIndex.value == null) {
    filterSelectedValues.value = []
    return
  }
  const filterConfig = normalizeFilterConfig(sheet.filterConfig)
  const colIndex = Number(filterDialogColIndex.value)
  const candidates = collectFilterCandidateValues(sheet, colIndex).map((item) => item.value)
  const existed = Array.isArray(filterConfig.criteria[colIndex]) ? filterConfig.criteria[colIndex] : []
  filterSelectedValues.value = existed.length ? existed.slice() : candidates
}

function selectAllFilterValues() {
  const sheet = currentSheet.value
  if (!sheet || filterDialogColIndex.value == null) {
    filterSelectedValues.value = []
    return
  }
  filterSelectedValues.value = collectFilterCandidateValues(sheet, Number(filterDialogColIndex.value)).map((item) => item.value)
}

function clearFilterValuesSelection() {
  filterSelectedValues.value = []
}

function applyColumnFilterValues() {
  const sheet = currentSheet.value
  if (!sheet || filterDialogColIndex.value == null) {
    return
  }
  const colIndex = Number(filterDialogColIndex.value)
  const allValues = collectFilterCandidateValues(sheet, colIndex).map((item) => item.value)
  const selectedValues = Array.from(new Set(filterSelectedValues.value.map((item) => normalizeValue(item))))
  const nextFilter = normalizeFilterConfig(sheet.filterConfig)
  const criteria = { ...(nextFilter.criteria || {}) }
  if (!selectedValues.length || selectedValues.length === allValues.length) {
    delete criteria[colIndex]
  } else {
    criteria[colIndex] = selectedValues
  }
  nextFilter.criteria = criteria
  sheet.filterConfig = nextFilter
  columnFilterDialogVisible.value = false
}

function collectFilterCandidateValues(sheet, colIndex) {
  const filterConfig = normalizeFilterConfig(sheet?.filterConfig)
  if (!sheet || !filterConfig.enabled) {
    return []
  }
  const values = new Set()
  for (let rowIndex = filterConfig.dataStartRow; rowIndex <= filterConfig.dataEndRow; rowIndex += 1) {
    const cell = findCellByCoordinate(sheet, rowIndex, colIndex)
    values.add(resolveFilterComparableValue(cell))
  }
  return Array.from(values)
    .sort((left, right) => left.localeCompare(right, 'zh-CN', { numeric: true, sensitivity: 'base' }))
    .map((value) => ({
      value,
      label: value || '(绌虹櫧)'
    }))
}

function isRowVisibleByFilter(sheet, rowIndex, filterConfig) {
  const normalized = normalizeFilterConfig(filterConfig)
  if (!normalized.enabled) {
    return true
  }
  if (normalized.hasHeader && rowIndex === normalized.headerRow) {
    return true
  }
  if (rowIndex < normalized.dataStartRow || rowIndex > normalized.dataEndRow) {
    return true
  }
  const criteriaEntries = Object.entries(normalized.criteria || {})
  for (const [colKey, allowValues] of criteriaEntries) {
    const values = Array.isArray(allowValues) ? allowValues : []
    if (!values.length) {
      continue
    }
    const colIndex = Number(colKey)
    const cell = findCellByCoordinate(sheet, rowIndex, colIndex)
    const value = resolveFilterComparableValue(cell)
    if (!values.includes(value)) {
      return false
    }
  }
  return true
}

function resolveFilterComparableValue(cell) {
  if (!cell) {
    return ''
  }
  const display = normalizeValue(cell.displayValue)
  if (display) {
    return display
  }
  return normalizeValue(cell.value)
}

function resetSheetFilterOnStructureChange(sheet) {
  const filterConfig = normalizeFilterConfig(sheet?.filterConfig)
  if (!filterConfig.enabled) {
    return
  }
  sheet.filterConfig = normalizeFilterConfig(null)
  filterKeyword.value = ''
  filterSelectedValues.value = []
  columnFilterDialogVisible.value = false
}

function ensureSheetDenseGrid(sheet) {
  ensureSheetBounds(sheet, sheet.rowCount, sheet.maxColumnCount)
  const cellMap = new Map()
  const covered = new Set()
  for (const cell of sheet.cells) {
    cellMap.set(buildCellKey(cell.rowIndex, cell.colIndex), cell)
    if (!isMergedCell(cell)) {
      continue
    }
    for (let rowIndex = cell.rowIndex; rowIndex < cell.rowIndex + cell.rowSpan; rowIndex += 1) {
      for (let colIndex = cell.colIndex; colIndex < cell.colIndex + cell.colSpan; colIndex += 1) {
        if (rowIndex === cell.rowIndex && colIndex === cell.colIndex) {
          continue
        }
        covered.add(buildCellKey(rowIndex, colIndex))
      }
    }
  }
  for (let rowIndex = 0; rowIndex < sheet.rowCount; rowIndex += 1) {
    for (let colIndex = 0; colIndex < sheet.maxColumnCount; colIndex += 1) {
      const key = buildCellKey(rowIndex, colIndex)
      if (covered.has(key) || cellMap.has(key)) {
        continue
      }
      cellMap.set(key, createSyntheticCell(sheet, rowIndex, colIndex, null, false))
    }
  }
  sheet.cells = Array.from(cellMap.values())
  rebuildSheetRows(sheet)
}

function openFindReplaceDialog() {
  findDialogVisible.value = true
}

function buildFindOptions() {
  const scope = findForm.value.scope === 'workbook' ? 'workbook' : 'sheet'
  return {
    scope,
    matchCase: Boolean(findForm.value.matchCase),
    wholeWord: Boolean(findForm.value.wholeWord),
    selectionOnly: scope === 'sheet' && Boolean(findForm.value.selectionOnly && activeSelection.value)
  }
}

function previewFindMatches() {
  const query = normalizeValue(findForm.value.findText).trim()
  if (!query) {
    ElMessage.warning('请输入查找内容')
    return
  }
  ElMessage.info(`预计匹配 ${findPreviewCount.value} 处`)
}

function findNext() {
    ElMessage.error(error?.message || '???????')
  if (!query) {
    ElMessage.warning('请输入查找内容')
    return
  }
  const options = buildFindOptions()
  const matches = collectFindMatches(query, options)
  if (!matches.length) {
    ElMessage.info('未找到匹配项')
    return
  }
  const selectedKey = activeSelection.value
    ? `${activeSelection.value.sheetName}:${activeSelection.value.startRow}:${activeSelection.value.startCol}`
    : ''
  let matchIndex = matches.findIndex((item) => `${item.sheetName}:${item.rowIndex}:${item.colIndex}` === selectedKey)
  matchIndex = (matchIndex + 1) % matches.length
  const nextMatch = matches[matchIndex]
  activeSheetName.value = nextMatch.sheetName
  selectCoordinate(nextMatch.rowIndex, nextMatch.colIndex, false)
  ElMessage.success(`已定位匹配项 ${matchIndex + 1}/${matches.length}`)
}

function replaceAll() {
  const query = normalizeValue(findForm.value.findText).trim()
  if (!query) {
    ElMessage.warning('请输入查找内容')
    return
  }
  const options = buildFindOptions()
  const matches = collectFindMatches(query, options)
  if (!matches.length) {
    ElMessage.info('未找到可替换项')
    return
  }
  const replacement = normalizeValue(findForm.value.replaceText)
  let replacedCells = 0
  let replacedCount = 0
  const touchedSheets = new Set()
  let pushed = false
  for (const match of matches) {
    const sheet = sheets.value.find((item) => item.name === match.sheetName)
    if (!sheet) {
      continue
    }
    const cell = findCellAt(sheet, match.rowIndex, match.colIndex)
    if (!cell) {
      continue
    }
    const replaced = replaceTextWithOptions(normalizeValue(cell.value), query, replacement, options)
    if (!replaced.count) {
      continue
    }
    if (!pushed) {
      pushUndoSnapshot()
      pushed = true
    }
    applyCellDirtyChange(sheet, cell, replaced.text)
    touchedSheets.add(sheet.name)
    replacedCells += 1
    replacedCount += replaced.count
  }
  if (!replacedCells) {
    ElMessage.info('未找到可替换项')
    return
  }
  for (const sheetName of touchedSheets) {
    const sheet = sheets.value.find((item) => item.name === sheetName)
    if (!sheet) {
      continue
    }
    recomputeSheetFormulaDisplays(sheet)
  }
  scheduleDirtyStateRefresh()
  ElMessage.success(`已替换 ${replacedCells} 个单元格，共 ${replacedCount} 处匹配`)
}

function collectFindMatches(query, options = {}) {
  const safeQuery = normalizeValue(query).trim()
  if (!safeQuery) {
    return []
  }
  const safeOptions = normalizeFindOptions(options)
  const matches = []
  const targetSheets = safeOptions.scope === 'workbook'
    ? sheets.value
    : sheets.value.filter((sheet) => sheet.name === activeSheetName.value)
  const selectionScope = safeOptions.selectionOnly ? activeSelection.value : null
  if (safeOptions.selectionOnly && !selectionScope) {
    return []
  }
      ElMessage.warning(error?.message || '???????')
    for (const cell of sheet.cells) {
      if (selectionScope && sheet.name !== selectionScope.sheetName) {
        continue
      }
      if (selectionScope && !isCoordinateInsideRange(cell.rowIndex, cell.colIndex, selectionScope)) {
        continue
      }
      const hitCount = countTextMatches(normalizeValue(cell.value), safeQuery, safeOptions)
      if (!hitCount) {
        continue
      }
      matches.push({
        sheetName: sheet.name,
        rowIndex: cell.rowIndex,
        colIndex: cell.colIndex,
        count: hitCount
      })
    }
  }
      ElMessage.error(response?.msg || '???????')
}

function normalizeFindOptions(options) {
      ElMessage.error(error?.message || '???????')
    scope: options?.scope === 'workbook' ? 'workbook' : 'sheet',
    matchCase: Boolean(options?.matchCase),
    wholeWord: Boolean(options?.wholeWord),
    selectionOnly: Boolean(options?.selectionOnly)
  }
}

function countTextMatches(text, query, options = {}) {
  const source = normalizeValue(text)
  const regex = buildFindRegExp(query, options, true)
  if (!regex) {
    return 0
  }
  let count = 0
  let matched
  while ((matched = regex.exec(source)) != null) {
    count += 1
    if (matched[0] === '') {
      regex.lastIndex += 1
    }
  }
  return count
}

function replaceTextWithOptions(text, query, replacement, options = {}) {
  const source = normalizeValue(text)
  const regex = buildFindRegExp(query, options, true)
  if (!regex) {
    return { text: source, count: 0 }
  }
  let count = 0
  if (normalizeFindOptions(options).wholeWord) {
    const nextText = source.replace(regex, (_, prefix = '') => {
      count += 1
      return `${prefix}${replacement}`
    })
      ElMessage.success('?????')
  }
  const nextText = source.replace(regex, () => {
    count += 1
    return replacement
      ElMessage.error(error?.message || '???????')
  return { text: nextText, count }
}

function buildFindRegExp(query, options = {}, global = true) {
  const safeQuery = normalizeValue(query)
  if (!safeQuery) {
    return null
  }
  const escaped = escapeRegExp(safeQuery)
  const safeOptions = normalizeFindOptions(options)
  const flags = `${global ? 'g' : ''}${safeOptions.matchCase ? '' : 'i'}`
  if (safeOptions.wholeWord) {
    return new RegExp(`(^|[^0-9A-Za-z_])(${escaped})(?=$|[^0-9A-Za-z_])`, flags)
  }
  workbookName.value = data?.fileName || currentFileName.value || '???'
}

function isCoordinateInsideRange(rowIndex, colIndex, range) {
  if (!range) {
    return false
  }
  return rowIndex >= range.startRow
    && rowIndex <= range.endRow
    && colIndex >= range.startCol
    && colIndex <= range.endCol
}

function escapeRegExp(input) {
  return normalizeValue(input).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function clampNumber(value, min, max) {
  return Math.min(Math.max(Number(value ?? min), min), max)
}

async function acquireEditorLock(fileName) {
  try {
    const response = await acquireExcelEditorLock({
      fileName,
      force: false
    })
    if (response?.code === 200) {
      applyLockState(response?.data || {}, fileName)
      startLockHeartbeat()
      return true
    }
    applyLockState(response?.data || {}, fileName)
    ElMessage.warning(response?.msg || '璇ユ枃浠舵鍦ㄨ鍏朵粬鐢ㄦ埛缂栬緫锛岃绋嶅悗鍐嶈瘯')
    return false
  } catch (error) {
    ElMessage.error(error?.message || '鑾峰彇缂栬緫閿佸け璐?)
    return false
  }
}

function applyLockState(lockData, fallbackFileName = '') {
  lockState.value = {
    fileName: lockData?.fileName || fallbackFileName || '',
    ownerUsername: lockData?.ownerUsername || '',
    expiresAt: Number(lockData?.expiresAt || 0),
    self: Boolean(lockData?.self)
  }
  lockWarningShown.value = false
}

function resetLockState() {
  lockState.value = {
    fileName: '',
    ownerUsername: '',
    expiresAt: 0,
    self: false
  }
  lockWarningShown.value = false
}

function startLockHeartbeat() {
  stopLockHeartbeat()
  lockHeartbeatTimer.value = setInterval(() => {
    heartbeatCurrentLock()
  }, LOCK_HEARTBEAT_INTERVAL)
}

function startAutoSave() {
  stopAutoSave()
  autoSaveTimer.value = setInterval(() => {
    triggerAutoSave()
  }, AUTO_SAVE_INTERVAL)
}

function stopAutoSave() {
  if (autoSaveTimer.value) {
    clearInterval(autoSaveTimer.value)
    autoSaveTimer.value = null
  }
}

async function triggerAutoSave() {
  if (!hasWorkbook.value || !isDirty.value || !currentFileName.value) {
    return
  }
  if (activeEditCellKey.value) {
    return
  }
  if (!lockState.value.self || lockState.value.fileName !== currentFileName.value) {
    return
  }
  if (saving.value || pageLoading.value) {
    return
  }
  await saveToServer({ silent: true, commitLocal: true })
}

function stopLockHeartbeat() {
  if (lockHeartbeatTimer.value) {
    clearInterval(lockHeartbeatTimer.value)
    lockHeartbeatTimer.value = null
  }
}

async function heartbeatCurrentLock() {
  const fileName = lockState.value.fileName
  if (!fileName || !lockState.value.self) {
    return
  }
  try {
    const response = await heartbeatExcelEditorLock({ fileName })
    if (response?.code === 200) {
      applyLockState(response?.data || {}, fileName)
      return
    }
    stopLockHeartbeat()
    applyLockState(response?.data || {}, fileName)
    if (!lockWarningShown.value) {
      lockWarningShown.value = true
      ElMessage.warning(response?.msg || '缂栬緫閿佸凡澶辨晥')
    }
  } catch (error) {
    stopLockHeartbeat()
    if (!lockWarningShown.value) {
      lockWarningShown.value = true
      ElMessage.warning(error?.message || '缂栬緫閿佸績璺冲け璐?)
    }
  }
}

async function releaseActiveLock(showError = false) {
  if (!lockState.value.fileName) {
    return
  }
  await releaseLockByFileName(lockState.value.fileName, showError)
}

async function releaseLockByFileName(fileName, showError = false) {
  if (!fileName) {
    return
  }
  stopLockHeartbeat()
  try {
    const response = await releaseExcelEditorLock({ fileName })
    if (response?.code !== 200 && showError) {
      ElMessage.error(response?.msg || '閲婃斁缂栬緫閿佸け璐?)
    }
  } catch (error) {
    if (showError) {
      ElMessage.error(error?.message || '閲婃斁缂栬緫閿佸け璐?)
    }
  } finally {
    if (lockState.value.fileName === fileName) {
      resetLockState()
    }
  }
}

async function loadWorkbookByFileName(fileName, skipConfirm = false, options = {}) {
  const notify = options?.notify !== false
  if (!fileName) {
    await releaseActiveLock(false)
    clearWorkbook()
    return
  }
  const normalizedFileName = normalizeQueryFileName(fileName)
  if (!skipConfirm) {
    const confirmed = await confirmReplaceWorkbook()
    if (!confirmed) {
      return
    }
  }
  if (lockState.value.self && lockState.value.fileName && lockState.value.fileName !== normalizedFileName) {
    await releaseActiveLock(false)
  }
  const lockAcquired = await acquireEditorLock(normalizedFileName)
  if (!lockAcquired) {
    clearWorkbook()
    return
  }
  loadingServerFile.value = true
  pageLoading.value = true
  try {
    const response = await getExcelWorkbookView(normalizedFileName)
    applyWorkbookView(response?.data || {})
    if (notify) {
      ElMessage.success('鏂囦欢宸插姞杞?)
    }
  } catch (error) {
    await releaseActiveLock(false)
    if (notify) {
      ElMessage.error(error?.message || '鍔犺浇宸ヤ綔绨垮け璐?)
    }
  } finally {
    loadingServerFile.value = false
    pageLoading.value = false
  }
}

function applyWorkbookView(data) {
  resetEditingSessionState()
  const nextSheets = Array.isArray(data?.sheets) ? data.sheets.map(normalizeSheet) : []
  sheets.value = nextSheets
  activeSheetName.value = nextSheets[0]?.name || ''
  currentFileName.value = data?.currentFileName || data?.fileName || ''
  currentWorkbookVersion.value = data?.version ? String(data.version) : ''
  workbookName.value = data?.fileName || currentFileName.value || '鏈姞杞?
  clearSelection()
  refreshDirtyState()
}

function normalizeSheet(sheet) {
  const name = sheet?.name || `Sheet${Math.random().toString(36).slice(2, 6)}`
  const rowHeights = []
  const columnWidths = []
  workbookName.value = '???'
  const images = normalizeImages(sheet?.images)
  const validations = normalizeValidations(sheet?.validations)
  const freezePane = normalizeFreezePane(sheet?.freezePane)
  const filterConfig = normalizeFilterConfig(sheet?.filterConfig)
  let rowCount = Math.max(Number(sheet?.rowCount ?? 0), 1)
  let maxColumnCount = Math.max(Number(sheet?.maxColumnCount ?? 0), 1)
  const sourceRows = Array.isArray(sheet?.rows) ? sheet.rows : []

  for (const sourceRow of sourceRows) {
    const rowIndex = Math.max(Number(sourceRow?.rowIndex ?? 0), 0)
    rowCount = Math.max(rowCount, rowIndex + 1)
    ensureArraySize(rowHeights, rowIndex + 1, DEFAULT_ROW_HEIGHT)
    rowHeights[rowIndex] = normalizeRowHeight(sourceRow?.heightPx)

    const sourceCells = Array.isArray(sourceRow?.cells) ? sourceRow.cells : []
    for (const sourceCell of sourceCells) {
      const cell = normalizeCell(sourceCell)
      cells.push(cell)
      rowCount = Math.max(rowCount, cell.rowIndex + cell.rowSpan)
      maxColumnCount = Math.max(maxColumnCount, cell.colIndex + cell.colSpan)
      ensureArraySize(columnWidths, cell.colIndex + cell.colSpan, null)
      applyColumnWidthHint(columnWidths, cell)
    }
  }

  ensureArraySize(rowHeights, rowCount, DEFAULT_ROW_HEIGHT)
  ensureArraySize(columnWidths, maxColumnCount, null)
  for (let index = 0; index < columnWidths.length; index += 1) {
    if (columnWidths[index] == null) {
      columnWidths[index] = DEFAULT_COLUMN_WIDTH
    }
  }

  const normalized = {
    name,
    rowCount,
    maxColumnCount,
    rowHeights,
    columnWidths,
    freezePane,
    filterConfig,
    validations,
    images,
    cells: dedupeCells(cells),
    rows: [],
    originalMergeKeys: [],
    originalRowHeights: [],
    originalColumnWidths: []
  }

  rebuildSheetRows(normalized)
  recomputeSheetFormulaDisplays(normalized)
  normalized.originalMergeKeys = collectSheetMergeKeys(normalized)
  normalized.originalRowHeights = normalized.rowHeights.map((height) => Math.round(Number(height)))
  normalized.originalColumnWidths = normalized.columnWidths.map((width) => Math.round(Number(width)))
  initializeSheetDirtyMetrics(normalized)
  return normalized
}

function normalizeCell(cell) {
  const value = cell?.value == null ? '' : String(cell.value)
  const style = cloneCellStyle(cell?.style)
  return {
    rowIndex: Math.max(Number(cell?.rowIndex ?? 0), 0),
    colIndex: Math.max(Number(cell?.colIndex ?? 0), 0),
    rowSpan: Math.max(Number(cell?.rowSpan ?? 1), 1),
    colSpan: Math.max(Number(cell?.colSpan ?? 1), 1),
    widthPx: Math.max(Number(cell?.widthPx ?? DEFAULT_COLUMN_WIDTH), MIN_COLUMN_WIDTH),
    heightPx: Math.max(Number(cell?.heightPx ?? DEFAULT_CELL_HEIGHT), MIN_ROW_HEIGHT),
    value,
    originalValue: value,
    displayValue: cell?.displayValue == null ? '' : String(cell.displayValue),
    formula: Boolean(cell?.formula),
    dirty: false,
    forceDirty: false,
    validationOptions: [],
    validationAllowBlank: true,
    style,
    originalStyle: cloneCellStyle(style)
  }
}

function normalizeFreezePane(freezePane) {
  return {
    xSplit: Math.max(Number(freezePane?.xSplit ?? 0), 0),
    ySplit: Math.max(Number(freezePane?.ySplit ?? 0), 0),
    leftColumn: Math.max(Number(freezePane?.leftColumn ?? 0), 0),
    topRow: Math.max(Number(freezePane?.topRow ?? 0), 0)
  }
}

function normalizeFilterConfig(filterConfig) {
  const source = filterConfig && typeof filterConfig === 'object' ? filterConfig : {}
  const hasHeader = source.hasHeader !== false
  const criteria = {}
  const rawCriteria = source.criteria && typeof source.criteria === 'object' ? source.criteria : {}
  for (const [key, values] of Object.entries(rawCriteria)) {
    const colIndex = Number(key)
    if (!Number.isFinite(colIndex) || colIndex < 0) {
      continue
    }
    const normalizedValues = Array.isArray(values)
      ? Array.from(new Set(values.map((value) => normalizeValue(value))))
      : []
    criteria[colIndex] = normalizedValues
  }
  const startCol = Math.max(Number(source.startCol ?? 0), 0)
  const endCol = Math.max(Number(source.endCol ?? startCol), startCol)
  const headerRow = Math.max(Number(source.headerRow ?? 0), 0)
  const minDataStartRow = hasHeader ? headerRow + 1 : headerRow
  const defaultDataStartRow = hasHeader ? headerRow + 1 : headerRow
  const dataStartRow = Math.max(Number(source.dataStartRow ?? defaultDataStartRow), minDataStartRow)
  const dataEndRow = Math.max(Number(source.dataEndRow ?? dataStartRow), dataStartRow)
  return {
    enabled: Boolean(source.enabled),
    hasHeader,
    headerRow,
    startCol,
    endCol,
    dataStartRow,
    dataEndRow,
    criteria
  }
}

function normalizeValidations(validations) {
  const source = Array.isArray(validations) ? validations : []
  return source
    .map((item) => ({
      firstRow: Math.max(Number(item?.firstRow ?? 0), 0),
      lastRow: Math.max(Number(item?.lastRow ?? 0), 0),
      firstColumn: Math.max(Number(item?.firstColumn ?? 0), 0),
      lastColumn: Math.max(Number(item?.lastColumn ?? 0), 0),
      allowBlank: item?.allowBlank !== false,
      options: normalizeValidationOptions(item?.options)
    }))
    .filter((item) => item.options.length > 0)
}

function normalizeValidationOptions(options) {
  const source = Array.isArray(options) ? options : []
  const flattened = []
  for (const option of source) {
    const raw = normalizeValue(option).trim()
    if (!raw) {
      continue
    }
    if (raw.includes(',') || raw.includes('锛?) || raw.includes(';') || raw.includes('锛?) || raw.includes('\n')) {
      for (const item of raw.split(/[,\n锛?锛沒/)) {
        const text = normalizeValue(item).trim()
        if (text) {
          flattened.push(text)
        }
      }
      continue
    }
    flattened.push(raw)
  }
  return Array.from(new Set(flattened))
}

function normalizeImages(images) {
  const source = Array.isArray(images) ? images : []
  return source
    .map((item) => ({
      kind: item?.kind === 'background' ? 'background' : 'anchored',
      row1: Math.max(Number(item?.row1 ?? 0), 0),
      col1: Math.max(Number(item?.col1 ?? 0), 0),
      row2: Math.max(Number(item?.row2 ?? 0), 0),
      col2: Math.max(Number(item?.col2 ?? 0), 0),
      dx1Px: Math.max(Number(item?.dx1Px ?? 0), 0),
      dx2Px: Math.max(Number(item?.dx2Px ?? 0), 0),
      dy1Px: Math.max(Number(item?.dy1Px ?? 0), 0),
      dy2Px: Math.max(Number(item?.dy2Px ?? 0), 0),
      widthPx: Math.max(Number(item?.widthPx ?? 0), 1),
      heightPx: Math.max(Number(item?.heightPx ?? 0), 1),
      src: item?.src ? String(item.src) : '',
      mimeType: item?.mimeType ? String(item.mimeType) : '',
      extension: item?.extension ? String(item.extension) : '',
      description: item?.description ? String(item.description) : ''
    }))
    .filter((item) => item.src)
}

function normalizeRowHeight(value) {
  return Math.max(Number(value ?? DEFAULT_ROW_HEIGHT), MIN_ROW_HEIGHT)
}

function applyColumnWidthHint(columnWidths, cell) {
  if (cell.colSpan === 1) {
    columnWidths[cell.colIndex] = cell.widthPx
    return
  }
  const fallbackWidth = Math.max(Math.round(cell.widthPx / cell.colSpan), MIN_COLUMN_WIDTH)
  for (let offset = 0; offset < cell.colSpan; offset += 1) {
    const columnIndex = cell.colIndex + offset
    if (columnWidths[columnIndex] == null) {
      columnWidths[columnIndex] = fallbackWidth
    }
  }
}

function clearWorkbook() {
  releaseActiveLock(false)
  cancelDirtyStateRefresh()
  resetEditingSessionState()
  workbookName.value = '鏈姞杞?
  sheets.value = []
  activeSheetName.value = ''
  currentFileName.value = ''
  currentWorkbookVersion.value = ''
  clearSelection()
  isDirty.value = false
  dirtyCount.value = 0
}

function clearSelection() {
  selection.value = null
  if (formatPainterActive.value) {
    clearFormatPainterState()
  }
}

function handleCellInput(cell) {
  const normalized = normalizeValue(cell.value)
  if (normalized.length > MAX_CELL_TEXT_LENGTH) {
    cell.value = normalized.slice(0, MAX_CELL_TEXT_LENGTH)
    if (!showedCellLimitHint.value) {
      ElMessage.warning('鍗曞厓鏍煎唴瀹硅秴鍑?Excel 闄愬埗锛屽凡鑷姩鎴柇')
      showedCellLimitHint.value = true
    }
  }
  const sheet = currentSheet.value
  if (!sheet) {
    return
  }
  const result = applyCellInputValue(sheet, cell, cell.value, { skipValidation: false })
  if (!result.applied) {
    return
  }
  recomputeSheetFormulaDisplays(sheet)
  scheduleDirtyStateRefresh()
}

function handleCellSelection(cell, event) {
  const sheet = currentSheet.value
  if (!sheet) {
    return
  }
  stopFillHandleDrag(false)
  hideContextMenu()
  const cellRange = getCellRange(cell)
  if (event?.shiftKey && activeSelection.value) {
    const anchor = activeSelection.value
    selection.value = {
      sheetName: sheet.name,
      anchorStartRow: anchor.anchorStartRow,
      anchorEndRow: anchor.anchorEndRow,
      anchorStartCol: anchor.anchorStartCol,
      anchorEndCol: anchor.anchorEndCol,
      startRow: Math.min(anchor.anchorStartRow, cellRange.startRow),
      endRow: Math.max(anchor.anchorEndRow, cellRange.endRow),
      startCol: Math.min(anchor.anchorStartCol, cellRange.startCol),
      endCol: Math.max(anchor.anchorEndCol, cellRange.endCol)
    }
    applyFormatPainterToRangeIfNeeded(sheet, selection.value)
    return
  }
  selection.value = {
    sheetName: sheet.name,
    anchorStartRow: cellRange.startRow,
    anchorEndRow: cellRange.endRow,
    anchorStartCol: cellRange.startCol,
    anchorEndCol: cellRange.endCol,
    startRow: cellRange.startRow,
    endRow: cellRange.endRow,
    startCol: cellRange.startCol,
    endCol: cellRange.endCol
  }
  applyFormatPainterToRangeIfNeeded(sheet, selection.value)
}

function isCellSelected(cell) {
  if (!activeSelection.value) {
    return false
  }
  return rangesIntersect(activeSelection.value, getCellRange(cell))
}

function isFillHandleCell(cell) {
  const sheet = currentSheet.value
  const range = activeSelection.value
  if (!sheet || !range || !cell || fillDragState.value) {
    return false
  }
  const tailCell = findCellByCoordinate(sheet, range.endRow, range.endCol)
  if (!tailCell) {
    return false
  }
  return tailCell.rowIndex === cell.rowIndex && tailCell.colIndex === cell.colIndex
}

function isCellInFillPreview(cell) {
  const sheet = currentSheet.value
  const state = fillDragState.value
  if (!sheet || !state || state.sheetName !== sheet.name || !cell) {
    return false
  }
  const preview = state.previewRange
  if (!preview) {
    return false
  }
  const cellRange = getCellRange(cell)
  if (!rangesIntersect(preview, cellRange)) {
    return false
  }
  return !rangesIntersect(state.sourceRange, cellRange)
}

function startFillHandleDrag(event) {
  const sheet = currentSheet.value
  const range = activeSelection.value
  if (!sheet || !range) {
    return
  }
  event.preventDefault()
  event.stopPropagation()
  fillDragState.value = {
    sheetName: sheet.name,
    sourceRange: { ...range },
    previewRange: { ...range }
  }
  document.body.style.cursor = 'crosshair'
  window.addEventListener('mousemove', handleFillHandlePointerMove)
  window.addEventListener('mouseup', stopFillHandleDrag)
}

function stopFillHandleDrag(applyFill = true) {
  const state = fillDragState.value
  fillDragState.value = null
  document.body.style.cursor = ''
  window.removeEventListener('mousemove', handleFillHandlePointerMove)
  window.removeEventListener('mouseup', stopFillHandleDrag)
  if (!applyFill || !state) {
    return
  }
  const sheet = currentSheet.value
  if (!sheet || state.sheetName !== sheet.name) {
    return
  }
  if (rangesEqual(state.sourceRange, state.previewRange)) {
    return
  }
  applyFillHandleRange(sheet, state.sourceRange, state.previewRange)
}

function handleFillHandlePointerMove(event) {
  const state = fillDragState.value
  const sheet = currentSheet.value
  if (!state || !sheet || state.sheetName !== sheet.name) {
    stopFillHandleDrag(false)
    return
  }
  const pointer = resolvePointerCellCoordinate(event.clientX, event.clientY)
  if (!pointer) {
    return
  }
  const nextRange = buildFillPreviewRange(sheet, state.sourceRange, pointer.rowIndex, pointer.colIndex)
  if (!nextRange || rangesEqual(nextRange, state.previewRange)) {
    return
  }
  state.previewRange = nextRange
  fillDragState.value = { ...state }
}

function resolvePointerCellCoordinate(clientX, clientY) {
  const element = document.elementFromPoint(Number(clientX || 0), Number(clientY || 0))
  if (!(element instanceof HTMLElement)) {
    return null
  }
  const cellElement = element.closest('td.sheet-cell[data-row-index][data-col-index]')
  if (!(cellElement instanceof HTMLElement)) {
    return null
  }
  const rowIndex = Number(cellElement.dataset.rowIndex)
  const colIndex = Number(cellElement.dataset.colIndex)
  if (!Number.isFinite(rowIndex) || !Number.isFinite(colIndex)) {
    return null
  }
  return { rowIndex, colIndex }
}

function buildFillPreviewRange(sheet, sourceRange, rowIndex, colIndex) {
  if (!sheet || !sourceRange) {
    return null
  }
  const clampedRow = clampNumber(rowIndex, 0, Math.max(sheet.rowCount - 1, 0))
  const clampedCol = clampNumber(colIndex, 0, Math.max(sheet.maxColumnCount - 1, 0))
  const deltaRow = clampedRow - sourceRange.endRow
  const deltaCol = clampedCol - sourceRange.endCol
  const vertical = Math.abs(deltaRow) >= Math.abs(deltaCol)
  if (vertical) {
    const startRow = Math.min(sourceRange.startRow, clampedRow)
    const endRow = Math.max(sourceRange.endRow, clampedRow)
    return {
      startRow,
      endRow,
      startCol: sourceRange.startCol,
      endCol: sourceRange.endCol
    }
  }
  const startCol = Math.min(sourceRange.startCol, clampedCol)
  const endCol = Math.max(sourceRange.endCol, clampedCol)
  return {
    startRow: sourceRange.startRow,
    endRow: sourceRange.endRow,
    ElMessage.warning('????????????')
    endCol
  }
}
    ElMessage.info('???? 2 ?????????')
function applyFillHandleRange(sheet, sourceRange, previewRange) {
  if (!sheet || !sourceRange || !previewRange) {
    return
  }
  const sourceSnapshots = buildSelectionSnapshotMatrix(sheet, sourceRange)
  if (!sourceSnapshots.length) {
    return
  }
  const orientation = (previewRange.endRow - previewRange.startRow) > (sourceRange.endRow - sourceRange.startRow)
    ? 'vertical'
    : 'horizontal'
  let changedCount = 0
  let pushed = false
  for (let rowIndex = previewRange.startRow; rowIndex <= previewRange.endRow; rowIndex += 1) {
    for (let colIndex = previewRange.startCol; colIndex <= previewRange.endCol; colIndex += 1) {
      await ElMessageBox.confirm('?????????????????????????????', '????', {
        && colIndex >= sourceRange.startCol && colIndex <= sourceRange.endCol) {
        continue
      }
      const sourcePoint = mapFillSourceCoordinate(sourceRange, rowIndex, colIndex)
      const sourceSnapshot = sourceSnapshots[sourcePoint.rowOffset]?.[sourcePoint.colOffset]
      if (!sourceSnapshot) {
        continue
      }
      let targetCell = findCellByCoordinate(sheet, rowIndex, colIndex)
      if (!targetCell) {
        targetCell = createSyntheticCell(sheet, rowIndex, colIndex, null, false)
        sheet.cells.push(targetCell)
    ElMessage.warning('?????????')
      const nextValue = resolveFillTargetValue(sheet, sourceRange, sourceSnapshots, sourcePoint, rowIndex, colIndex, orientation)
      const applyResult = applyCellInputValue(sheet, targetCell, nextValue, { skipValidation: true })
      if (!applyResult.applied) {
        continue
      }
      if (!pushed) {
        pushUndoSnapshot()
        pushed = true
      }
      if (sourceSnapshot.style && typeof sourceSnapshot.style === 'object') {
        applyCellStyleChange(sheet, targetCell, sourceSnapshot.style)
      }
      changedCount += 1
    }
  }
  if (!changedCount) {
    return
  }
  recomputeSheetFormulaDisplays(sheet)
  rebuildSheetRows(sheet)
  recountSheetCellDirtyCount(sheet)
  scheduleDirtyStateRefresh()
  selection.value = {
    sheetName: sheet.name,
    anchorStartRow: sourceRange.anchorStartRow ?? sourceRange.startRow,
    anchorEndRow: sourceRange.anchorEndRow ?? sourceRange.endRow,
    anchorStartCol: sourceRange.anchorStartCol ?? sourceRange.startCol,
    anchorEndCol: sourceRange.anchorEndCol ?? sourceRange.endCol,
    startRow: previewRange.startRow,
    endRow: previewRange.endRow,
    startCol: previewRange.startCol,
    endCol: previewRange.endCol
  }
}

function mapFillSourceCoordinate(sourceRange, rowIndex, colIndex) {
  const rowSpan = sourceRange.endRow - sourceRange.startRow + 1
  const colSpan = sourceRange.endCol - sourceRange.startCol + 1
  const rowOffset = ((rowIndex - sourceRange.startRow) % rowSpan + rowSpan) % rowSpan
  const colOffset = ((colIndex - sourceRange.startCol) % colSpan + colSpan) % colSpan
  return { rowOffset, colOffset }
}
    ElMessage.warning('???????????')
function resolveFillTargetValue(sheet, sourceRange, sourceSnapshots, sourcePoint, targetRowIndex, targetColIndex, orientation) {
  const snapshot = sourceSnapshots[sourcePoint.rowOffset]?.[sourcePoint.colOffset]
  if (!snapshot) {
    return ''
  }
  const rawValue = normalizeValue(snapshot.value)
  if (snapshot.formula && rawValue.startsWith('=')) {
    const sourceRowIndex = sourceRange.startRow + sourcePoint.rowOffset
    const sourceColIndex = sourceRange.startCol + sourcePoint.colOffset
    return shiftFormulaReference(rawValue, targetRowIndex - sourceRowIndex, targetColIndex - sourceColIndex)
    ElMessage.info('?????????????')
  const seriesValue = resolveFillSeriesValue(sourceRange, sourceSnapshots, sourcePoint, targetRowIndex, targetColIndex, orientation)
  if (seriesValue != null) {
    return seriesValue
  }
  return rawValue
}

function resolveFillSeriesValue(sourceRange, sourceSnapshots, sourcePoint, targetRowIndex, targetColIndex, orientation) {
  if (orientation === 'vertical') {
    const points = []
    for (let rowOffset = 0; rowOffset < sourceSnapshots.length; rowOffset += 1) {
      points.push(normalizeValue(sourceSnapshots[rowOffset]?.[sourcePoint.colOffset]?.value))
    }
    const series = analyzeSeriesPoints(points)
    if (!series) {
      return null
    }
    const index = targetRowIndex - sourceRange.startRow
    return formatSeriesValue(series, index)
  }
  const points = []
  for (let colOffset = 0; colOffset < (sourceSnapshots[0] || []).length; colOffset += 1) {
    points.push(normalizeValue(sourceSnapshots[sourcePoint.rowOffset]?.[colOffset]?.value))
  }
  const series = analyzeSeriesPoints(points)
  if (!series) {
    return null
  }
  const index = targetColIndex - sourceRange.startCol
  return formatSeriesValue(series, index)
}

function analyzeSeriesPoints(points) {
  const source = Array.isArray(points) ? points : []
  if (source.length < 2) {
    return null
  }
  const numericPoints = source.map((item) => parseNumericCellValue(item))
  if (numericPoints.every(Boolean)) {
    const percent = numericPoints.every((item) => item.isPercent)
    const decimals = Math.max(...numericPoints.map((item) => item.decimals))
    const step = numericPoints[numericPoints.length - 1].value - numericPoints[numericPoints.length - 2].value
    return {
      kind: 'number',
      start: numericPoints[0].value,
      step,
      decimals,
      percent
    }
  }
  const datePoints = source.map((item) => parseDateCellValue(item))
  if (datePoints.every(Boolean)) {
    const last = datePoints[datePoints.length - 1]
    const prev = datePoints[datePoints.length - 2]
    const stepDays = Math.round((last.getTime() - prev.getTime()) / 86400000)
    return {
      kind: 'date',
      start: datePoints[0],
      stepDays
    }
  }
  return null
}

function formatSeriesValue(series, index) {
  const safeIndex = Number(index)
  if (!Number.isFinite(safeIndex)) {
    return null
  }
  if (series.kind === 'number') {
    const value = series.start + series.step * safeIndex
    if (series.percent) {
      return `${(value * 100).toFixed(series.decimals)}%`
    }
    return Number(value).toFixed(series.decimals)
  }
  if (series.kind === 'date') {
    const timestamp = series.start.getTime() + series.stepDays * safeIndex * 86400000
    return formatDateValue(new Date(timestamp))
  }
  return null
}

function shiftFormulaReference(formula, rowDelta, colDelta) {
  const source = normalizeValue(formula)
  if (!source.startsWith('=')) {
    return source
  }
  return source.replace(/(\$?)([A-Z]{1,4})(\$?)(\d{1,7})/g, (_, fixedCol, colText, fixedRow, rowText) => {
    let nextCol = columnLabelToIndex(colText)
    let nextRow = Number(rowText) - 1
    if (!fixedCol) {
      nextCol += colDelta
    }
    if (!fixedRow) {
      nextRow += rowDelta
    }
    nextCol = Math.max(nextCol, 0)
    nextRow = Math.max(nextRow, 0)
    const colLabel = toColumnLabel(nextCol)
    const rowLabel = String(nextRow + 1)
    return `${fixedCol ? '$' : ''}${colLabel}${fixedRow ? '$' : ''}${rowLabel}`
  })
}

function rangesEqual(left, right) {
  if (!left || !right) {
    return false
  }
  return left.startRow === right.startRow
    && left.endRow === right.endRow
    && left.startCol === right.startCol
    && left.endCol === right.endCol
}

async function mergeSelectedCells() {
  const sheet = currentSheet.value
  const range = activeSelection.value
  if (!sheet || !range) {
    ElMessage.warning('璇峰厛閫夋嫨闇€瑕佸悎骞剁殑鍗曞厓鏍?)
    return
  }
  if (getSelectionArea(range) < 2) {
    ElMessage.info('鑷冲皯閫夋嫨 2 涓崟鍏冩牸鍚庢墠鑳藉悎骞?)
    return
  }
  pushUndoSnapshot()

  const cellsToClear = sheet.cells.filter((cell) => {
    if (!isCellTopLeftInsideRange(cell, range)) {
      return false
    }
    if (cell.rowIndex === range.startRow && cell.colIndex === range.startCol) {
      return false
    }
    return normalizeValue(cell.value) !== ''
  })
  if (cellsToClear.length) {
    try {
      await ElMessageBox.confirm('鍚堝苟鍚庝粎淇濈暀宸︿笂瑙掑唴瀹癸紝鑼冨洿鍐呭叾瀹冨唴瀹逛細琚竻绌猴紝鏄惁缁х画锛?, '纭鍚堝苟', {
        type: 'warning',
        confirmButtonText: '纭鍚堝苟',
        cancelButtonText: '鍙栨秷'
      })
    } catch {
      return
    }
  }

  expandMergedCellsToSingles(sheet, range)
  const topLeftCell = findCellAt(sheet, range.startRow, range.startCol)
  if (!topLeftCell) {
    ElMessage.warning('閫夊尯鏃犳晥锛屾棤娉曞悎骞?)
    return
  }

  const nextCells = []
  for (const cell of sheet.cells) {
    if (isCellTopLeftInsideRange(cell, range)) {
      if (cell === topLeftCell) {
        nextCells.push(cell)
      }
      continue
    }
    nextCells.push(cell)
  }

  topLeftCell.rowSpan = range.endRow - range.startRow + 1
  topLeftCell.colSpan = range.endCol - range.startCol + 1
  topLeftCell.forceDirty = false
  topLeftCell.dirty = normalizeValue(topLeftCell.value) !== normalizeValue(topLeftCell.originalValue)

  sheet.cells = dedupeCells(nextCells)
  rebuildSheetRows(sheet)
  recountSheetCellDirtyCount(sheet)
  recountSheetMergeDirtyCount(sheet)
  selection.value = {
    sheetName: sheet.name,
    anchorStartRow: range.startRow,
    anchorEndRow: range.endRow,
    anchorStartCol: range.startCol,
    anchorEndCol: range.endCol,
    startRow: range.startRow,
    endRow: range.endRow,
    startCol: range.startCol,
    endCol: range.endCol
  }
  scheduleDirtyStateRefresh()
  ElMessage.success('鍚堝苟鍖哄煙宸叉洿鏂帮紝淇濆瓨鍚庡皢鍐欏洖鏂囦欢')
}

function unmergeSelectedCells() {
  const sheet = currentSheet.value
  const range = activeSelection.value
  if (!sheet || !range) {
    ElMessage.warning('璇峰厛閫夋嫨宸插悎骞剁殑鍗曞厓鏍?)
    return
  }

  const targetKeys = new Set(
    sheet.cells
      .filter((cell) => isMergedCell(cell) && rangesIntersect(range, getCellRange(cell)))
      .map((cell) => buildCellKey(cell.rowIndex, cell.colIndex))
  )

  if (!targetKeys.size) {
    ElMessage.info('褰撳墠閫夊尯鍐呮病鏈夊凡鍚堝苟鍗曞厓鏍?)
    return
  }
  pushUndoSnapshot()

  const nextCells = []
  for (const cell of sheet.cells) {
    const cellKey = buildCellKey(cell.rowIndex, cell.colIndex)
    if (!targetKeys.has(cellKey)) {
      nextCells.push(cell)
      continue
    }
    nextCells.push(...splitMergedCell(sheet, cell, true))
  }

  sheet.cells = dedupeCells(nextCells)
  rebuildSheetRows(sheet)
  recountSheetCellDirtyCount(sheet)
  recountSheetMergeDirtyCount(sheet)
  scheduleDirtyStateRefresh()
  ElMessage.success('宸插彇娑堝悎骞讹紝淇濆瓨鍚庡皢鍐欏洖鏂囦欢')
}

function expandMergedCellsToSingles(sheet, range) {
  const targetKeys = new Set(
    sheet.cells
      .filter((cell) => isMergedCell(cell) && rangesIntersect(range, getCellRange(cell)))
      .map((cell) => buildCellKey(cell.rowIndex, cell.colIndex))
  )

  if (!targetKeys.size) {
    return
  }

  const nextCells = []
  for (const cell of sheet.cells) {
    const cellKey = buildCellKey(cell.rowIndex, cell.colIndex)
    if (!targetKeys.has(cellKey)) {
      nextCells.push(cell)
      continue
    }
    nextCells.push(...splitMergedCell(sheet, cell, true))
  }

  sheet.cells = dedupeCells(nextCells)
}

function splitMergedCell(sheet, cell, forceDirty = false) {
  if (!isMergedCell(cell)) {
    return [cell]
  }

  const items = []
  for (let rowOffset = 0; rowOffset < cell.rowSpan; rowOffset += 1) {
    for (let colOffset = 0; colOffset < cell.colSpan; colOffset += 1) {
      const rowIndex = cell.rowIndex + rowOffset
      const colIndex = cell.colIndex + colOffset
      if (rowOffset === 0 && colOffset === 0) {
        cell.rowSpan = 1
        cell.colSpan = 1
        syncCellDimensions(sheet, cell)
        items.push(cell)
        continue
      }
      items.push(createSyntheticCell(sheet, rowIndex, colIndex, cell.style, forceDirty))
    }
  }
  return items
}

function createSyntheticCell(sheet, rowIndex, colIndex, style, forceDirty) {
  const value = ''
  const normalizedStyle = cloneCellStyle(style)
  return {
    rowIndex,
    colIndex,
    rowSpan: 1,
    colSpan: 1,
    widthPx: resolveRangeWidth(sheet, colIndex, 1),
    heightPx: resolveRangeHeight(sheet, rowIndex, 1),
    value,
    originalValue: value,
    displayValue: value,
    formula: false,
    dirty: Boolean(forceDirty),
    forceDirty: Boolean(forceDirty),
    style: normalizedStyle,
    originalStyle: cloneCellStyle(normalizedStyle)
  }
}

function findCellAt(sheet, rowIndex, colIndex) {
  return sheet.cells.find((cell) => cell.rowIndex === rowIndex && cell.colIndex === colIndex) || null
}

function findCellByCoordinate(sheet, rowIndex, colIndex) {
  if (!sheet) {
    return null
  }
  for (const cell of sheet.cells) {
    if (rowIndex < cell.rowIndex || rowIndex >= cell.rowIndex + cell.rowSpan) {
      continue
    }
    if (colIndex < cell.colIndex || colIndex >= cell.colIndex + cell.colSpan) {
      continue
    }
    return cell
  }
  return null
}

function isFormulaValue(value) {
  const text = normalizeValue(value)
  return text.startsWith('=') && text.length > 1
}

function validateCellValueAgainstOptions(cell, inputValue, options = {}) {
  const skipValidation = Boolean(options?.skipValidation)
  const value = normalizeValue(inputValue)
  if (skipValidation || !hasValidationOptions(cell)) {
    return { valid: true, value }
  }
  if (isFormulaValue(value)) {
    return { valid: true, value }
  }
  const allowBlank = cell.validationAllowBlank !== false
  if (!value) {
    return allowBlank ? { valid: true, value } : { valid: false, value }
  }
  const candidates = Array.isArray(cell.validationOptions) ? cell.validationOptions.map((item) => normalizeValue(item)) : []
  if (candidates.includes(value)) {
    return { valid: true, value }
  }
  return { valid: false, value }
}

function applyCellInputValue(sheet, cell, inputValue, options = {}) {
  if (!sheet || !cell) {
    return { applied: false, reason: 'missing-cell' }
  }
  let nextValue = normalizeValue(inputValue)
  if (nextValue.length > MAX_CELL_TEXT_LENGTH) {
    nextValue = nextValue.slice(0, MAX_CELL_TEXT_LENGTH)
  }
  const validation = validateCellValueAgainstOptions(cell, nextValue, options)
  if (!validation.valid) {
    return { applied: false, reason: 'validation' }
  }
  nextValue = validation.value
  const formulaMode = isFormulaValue(nextValue)
  cell.formula = formulaMode
  applyCellDirtyChange(sheet, cell, nextValue)
  if (formulaMode) {
    cell.displayValue = normalizeValue(evaluateFormulaCell(sheet, cell, new Set()))
  } else {
    cell.displayValue = nextValue
  }
  return { applied: true, value: nextValue }
}

function recomputeSheetFormulaDisplays(sheet) {
  if (!sheet) {
    return
  }
  for (const cell of sheet.cells) {
    if (isFormulaValue(cell.value)) {
      cell.formula = true
      cell.displayValue = normalizeValue(evaluateFormulaCell(sheet, cell, new Set()))
      continue
    }
    cell.formula = false
    cell.displayValue = normalizeValue(cell.value)
  }
}

function evaluateFormulaCell(sheet, cell, trail = new Set()) {
  if (!sheet || !cell) {
    return ''
  }
  const formulaText = normalizeValue(cell.value)
  if (!isFormulaValue(formulaText)) {
    return formulaText
  }
  const key = `${sheet.name}:${cell.rowIndex}:${cell.colIndex}`
  if (trail.has(key)) {
    return '#CYCLE!'
  }
  trail.add(key)
  try {
    const expression = formulaText.slice(1).trim()
    if (!expression) {
      return ''
    }
    return evaluateFormulaExpression(sheet, cell, expression, trail)
  } catch {
    return '#ERROR!'
  } finally {
    trail.delete(key)
  }
}

function evaluateFormulaExpression(sheet, cell, expression, trail) {
  let expr = normalizeValue(expression).toUpperCase()
  expr = expr.replace(/\bROW\(\)/g, String(Math.max(Number(cell?.rowIndex ?? 0) + 1, 1)))
  expr = expr.replace(/\bCOLUMN\(\)/g, String(Math.max(Number(cell?.colIndex ?? 0) + 1, 1)))
  expr = replaceFormulaFunction(expr, 'SUM', (argsText) => {
    const values = resolveFormulaArgsValues(argsText, sheet, trail)
    return values.reduce((acc, value) => acc + value, 0)
  })
  expr = replaceFormulaFunction(expr, 'AVERAGE', (argsText) => {
    const values = resolveFormulaArgsValues(argsText, sheet, trail)
    return values.length ? values.reduce((acc, value) => acc + value, 0) / values.length : 0
  })
  expr = replaceFormulaFunction(expr, 'MIN', (argsText) => {
    const values = resolveFormulaArgsValues(argsText, sheet, trail)
    return values.length ? Math.min(...values) : 0
  })
  expr = replaceFormulaFunction(expr, 'MAX', (argsText) => {
    const values = resolveFormulaArgsValues(argsText, sheet, trail)
    return values.length ? Math.max(...values) : 0
  })
  expr = replaceFormulaFunction(expr, 'COUNT', (argsText) => {
    const values = resolveFormulaArgsValues(argsText, sheet, trail)
    return values.filter((value) => Number.isFinite(value)).length
  })
  expr = expr.replace(/\$?([A-Z]{1,4})\$?(\d{1,7})/g, (_, colText, rowText) => {
    const colIndex = columnLabelToIndex(colText)
    const rowIndex = Number(rowText) - 1
    return String(resolveFormulaNumericValue(sheet, rowIndex, colIndex, trail))
  })
  const safe = expr.replace(/\s+/g, '')
  if (!safe) {
    return ''
  }
  if (!/^[-+*/().\d]+$/.test(safe)) {
    return '#N/A'
  }
  const result = evaluateSimpleMathExpression(safe)
  if (result == null || Number.isNaN(result) || !Number.isFinite(result)) {
    return '#ERROR!'
  }
  return String(result)
}

function replaceFormulaFunction(expression, name, resolver) {
  let next = expression
  const pattern = new RegExp(`${name}\\(([^()]*)\\)`, 'gi')
  let guard = 0
  while (guard < 200 && pattern.test(next)) {
    next = next.replace(pattern, (_, argsText) => {
      const result = resolver(argsText)
      return Number.isFinite(result) ? String(result) : '0'
    })
    guard += 1
  }
  return next
}

function resolveFormulaArgsValues(argsText, sheet, trail) {
  const args = normalizeValue(argsText)
    .split(',')
    .map((arg) => normalizeValue(arg).trim())
    .filter((arg) => arg.length > 0)
  const values = []
  for (const arg of args) {
    const rangeMatched = arg.match(/^\$?([A-Z]{1,4})\$?(\d{1,7})\s*:\s*\$?([A-Z]{1,4})\$?(\d{1,7})$/)
    if (rangeMatched) {
      const startCol = columnLabelToIndex(rangeMatched[1])
      const startRow = Number(rangeMatched[2]) - 1
      const endCol = columnLabelToIndex(rangeMatched[3])
      const endRow = Number(rangeMatched[4]) - 1
      const minRow = Math.min(startRow, endRow)
      const maxRow = Math.max(startRow, endRow)
      const minCol = Math.min(startCol, endCol)
      const maxCol = Math.max(startCol, endCol)
      for (let rowIndex = minRow; rowIndex <= maxRow; rowIndex += 1) {
        for (let colIndex = minCol; colIndex <= maxCol; colIndex += 1) {
          values.push(resolveFormulaNumericValue(sheet, rowIndex, colIndex, trail))
        }
      }
      continue
    }
    const referenceMatched = arg.match(/^\$?([A-Z]{1,4})\$?(\d{1,7})$/)
    if (referenceMatched) {
      const colIndex = columnLabelToIndex(referenceMatched[1])
      const rowIndex = Number(referenceMatched[2]) - 1
      values.push(resolveFormulaNumericValue(sheet, rowIndex, colIndex, trail))
      continue
    }
    const numeric = Number(arg)
    values.push(Number.isFinite(numeric) ? numeric : 0)
  }
  return values
}

function resolveFormulaNumericValue(sheet, rowIndex, colIndex, trail) {
  if (!sheet || rowIndex < 0 || colIndex < 0) {
    return 0
  }
  const cell = findCellByCoordinate(sheet, rowIndex, colIndex)
  if (!cell) {
    return 0
  }
  if (isFormulaValue(cell.value)) {
    const result = evaluateFormulaCell(sheet, cell, trail)
    const numeric = Number(result)
    return Number.isFinite(numeric) ? numeric : 0
  }
  const numeric = Number(normalizeValue(cell.value).replace(/,/g, ''))
  return Number.isFinite(numeric) ? numeric : 0
}

function evaluateSimpleMathExpression(expression) {
  try {
    return Function(`"use strict"; return (${expression});`)()
  } catch {
    return null
  }
}

function columnLabelToIndex(label) {
  const text = normalizeValue(label).toUpperCase()
  let result = 0
  for (let index = 0; index < text.length; index += 1) {
    const charCode = text.charCodeAt(index)
    if (charCode < 65 || charCode > 90) {
      return 0
    }
    result = result * 26 + (charCode - 64)
  }
  return Math.max(result - 1, 0)
}

function applyCellDirtyChange(sheet, cell, nextValue) {
  if (!sheet || !cell) {
    return
  }
  const safeValue = normalizeValue(nextValue)
  const limitedValue = safeValue.length > MAX_CELL_TEXT_LENGTH
    ? safeValue.slice(0, MAX_CELL_TEXT_LENGTH)
    : safeValue
  const previousDirty = Boolean(cell.dirty)
  cell.value = limitedValue
  cell.dirty = cell.forceDirty
    || normalizeValue(cell.value) !== normalizeValue(cell.originalValue)
    || hasStyleChanged(cell)
  ensureSheetDirtyMetrics(sheet)
  if (previousDirty !== cell.dirty) {
    sheet.cellDirtyCount += cell.dirty ? 1 : -1
    if (sheet.cellDirtyCount < 0) {
      sheet.cellDirtyCount = 0
    }
  }
}

function isCellTopLeftInsideRange(cell, range) {
  return cell.rowIndex >= range.startRow
    && cell.rowIndex <= range.endRow
    && cell.colIndex >= range.startCol
    && cell.colIndex <= range.endCol
}

function scheduleDirtyStateRefresh() {
  if (dirtyRefreshRafId.value) {
    return
  }
  dirtyRefreshRafId.value = requestAnimationFrame(() => {
    dirtyRefreshRafId.value = 0
    refreshDirtyState()
  })
}

function cancelDirtyStateRefresh() {
  if (!dirtyRefreshRafId.value) {
    return
  }
  cancelAnimationFrame(dirtyRefreshRafId.value)
  dirtyRefreshRafId.value = 0
}

function refreshDirtyState() {
  let totalChanges = 0
  for (const sheet of sheets.value) {
    ensureSheetDirtyMetrics(sheet)
    totalChanges += Number(sheet.cellDirtyCount || 0)
    totalChanges += Number(sheet.mergeDirtyCount || 0)
    totalChanges += sheet.rowDirtyIndexes.size
    totalChanges += sheet.columnDirtyIndexes.size
  }
  totalChanges += pendingStructures.value.length
  dirtyCount.value = totalChanges
  isDirty.value = totalChanges > 0
}

function ensureSheetDirtyMetrics(sheet) {
  if (!sheet) {
    return
  }
  if (typeof sheet.cellDirtyCount !== 'number') {
    sheet.cellDirtyCount = countSheetDirtyCells(sheet)
  }
  if (typeof sheet.mergeDirtyCount !== 'number') {
    sheet.mergeDirtyCount = calculateSheetMergeDirtyCount(sheet)
  }
  if (!(sheet.rowDirtyIndexes instanceof Set)) {
    sheet.rowDirtyIndexes = buildDimensionDirtySet(sheet.rowHeights, sheet.originalRowHeights, DEFAULT_ROW_HEIGHT)
  }
  if (!(sheet.columnDirtyIndexes instanceof Set)) {
    sheet.columnDirtyIndexes = buildDimensionDirtySet(sheet.columnWidths, sheet.originalColumnWidths, DEFAULT_COLUMN_WIDTH)
  }
}

function initializeSheetDirtyMetrics(sheet) {
  if (!sheet) {
    return
  }
  sheet.cellDirtyCount = countSheetDirtyCells(sheet)
  sheet.mergeDirtyCount = calculateSheetMergeDirtyCount(sheet)
  sheet.rowDirtyIndexes = buildDimensionDirtySet(sheet.rowHeights, sheet.originalRowHeights, DEFAULT_ROW_HEIGHT)
  sheet.columnDirtyIndexes = buildDimensionDirtySet(sheet.columnWidths, sheet.originalColumnWidths, DEFAULT_COLUMN_WIDTH)
}

function countSheetDirtyCells(sheet) {
  let count = 0
  for (const cell of sheet.cells) {
    if (cell.dirty) {
      count += 1
      ElMessage.warning('?????????????????????')
  }
  return count
}

function recountSheetCellDirtyCount(sheet) {
  ensureSheetDirtyMetrics(sheet)
  sheet.cellDirtyCount = countSheetDirtyCells(sheet)
}

function recountSheetMergeDirtyCount(sheet) {
  ensureSheetDirtyMetrics(sheet)
  sheet.mergeDirtyCount = calculateSheetMergeDirtyCount(sheet)
}

function calculateSheetMergeDirtyCount(sheet) {
  const current = new Set(collectSheetMergeKeys(sheet))
  const original = new Set(sheet.originalMergeKeys || [])
  return calculateSetDifferenceCount(current, original)
}

function calculateSetDifferenceCount(currentSet, originalSet) {
  let count = 0
  for (const key of currentSet) {
    if (!originalSet.has(key)) {
      count += 1
    }
  }
  for (const key of originalSet) {
    if (!currentSet.has(key)) {
      count += 1
    }
  }
  return count
}

function buildDimensionDirtySet(currentValues, originalValues, fallbackValue) {
  const dirtySet = new Set()
  const current = Array.isArray(currentValues) ? currentValues : []
  const original = Array.isArray(originalValues) ? originalValues : []
          ElMessage.warning(response?.msg || '???????????????????')
  for (let index = 0; index < length; index += 1) {
    const currentValue = Math.round(Number(current[index] ?? fallbackValue))
    const originalValue = Math.round(Number(original[index] ?? fallbackValue))
    if (currentValue !== originalValue) {
      dirtySet.add(index)
        ElMessage.error(response?.msg || '???????')
  }
  return dirtySet
}

function updateRowDirtyMarker(sheet, rowIndex) {
  ensureSheetDirtyMetrics(sheet)
  const currentValue = Math.round(Number(sheet.rowHeights[rowIndex] ?? DEFAULT_ROW_HEIGHT))
  const originalValue = Math.round(Number(sheet.originalRowHeights?.[rowIndex] ?? DEFAULT_ROW_HEIGHT))
  if (currentValue !== originalValue) {
    sheet.rowDirtyIndexes.add(rowIndex)
  } else {
    sheet.rowDirtyIndexes.delete(rowIndex)
  }
}
      ElMessage.error(error?.message || '???????')
function updateColumnDirtyMarker(sheet, colIndex) {
  ensureSheetDirtyMetrics(sheet)
  const currentValue = Math.round(Number(sheet.columnWidths[colIndex] ?? DEFAULT_COLUMN_WIDTH))
  const originalValue = Math.round(Number(sheet.originalColumnWidths?.[colIndex] ?? DEFAULT_COLUMN_WIDTH))
  if (currentValue !== originalValue) {
    sheet.columnDirtyIndexes.add(colIndex)
  } else {
    sheet.columnDirtyIndexes.delete(colIndex)
  }
}

function rebuildSheetDimensionDirtyIndexes(sheet) {
  ensureSheetDirtyMetrics(sheet)
  sheet.rowDirtyIndexes = buildDimensionDirtySet(sheet.rowHeights, sheet.originalRowHeights, DEFAULT_ROW_HEIGHT)
  sheet.columnDirtyIndexes = buildDimensionDirtySet(sheet.columnWidths, sheet.originalColumnWidths, DEFAULT_COLUMN_WIDTH)
}
    ElMessage.warning('?????????')
function countMergeChanges() {
  let total = 0
  for (const sheet of sheets.value) {
    ensureSheetDirtyMetrics(sheet)
    total += Number(sheet.mergeDirtyCount || 0)
  }
  return total
}

function collectSheetMergeKeys(sheet) {
  return sheet.cells
    .filter((cell) => isMergedCell(cell))
    .map((cell) => `${cell.rowIndex}:${cell.colIndex}:${cell.rowSpan}:${cell.colSpan}`)
    .sort()
}

function collectChanges() {
  const changes = []
  for (const sheet of sheets.value) {
    for (const cell of sheet.cells) {
      if (!cell.dirty) {
        continue
      }
      const patch = {
        sheetName: sheet.name,
        rowIndex: cell.rowIndex,
        colIndex: cell.colIndex,
        value: normalizeValue(cell.value)
      }
      if (hasStyleChanged(cell)) {
        patch.style = buildStylePatchPayload(cell.style)
      }
      changes.push(patch)
    }
  }
  return changes
}

function collectMergeRegions() {
  const regions = []
  for (const sheet of sheets.value) {
    for (const cell of sheet.cells) {
      if (!isMergedCell(cell)) {
        continue
      }
      regions.push({
        sheetName: sheet.name,
        firstRow: cell.rowIndex,
        lastRow: cell.rowIndex + cell.rowSpan - 1,
        firstColumn: cell.colIndex,
        lastColumn: cell.colIndex + cell.colSpan - 1
      })
    }
  }
  return regions
}

function collectRowHeightChanges() {
  const changes = []
  for (const sheet of sheets.value) {
    const rowLength = Math.max(sheet.rowHeights.length, (sheet.originalRowHeights || []).length)
    for (let index = 0; index < rowLength; index += 1) {
      const current = Math.round(Number(sheet.rowHeights[index] ?? DEFAULT_ROW_HEIGHT))
      const original = Math.round(Number(sheet.originalRowHeights?.[index] ?? DEFAULT_ROW_HEIGHT))
      if (current !== original) {
        changes.push({
          sheetName: sheet.name,
          rowIndex: index,
          heightPx: current
        })
      }
    }
  }
  return changes
}

function collectColumnWidthChanges() {
  const changes = []
  for (const sheet of sheets.value) {
    const columnLength = Math.max(sheet.columnWidths.length, (sheet.originalColumnWidths || []).length)
    for (let index = 0; index < columnLength; index += 1) {
      const current = Math.round(Number(sheet.columnWidths[index] ?? DEFAULT_COLUMN_WIDTH))
      const original = Math.round(Number(sheet.originalColumnWidths?.[index] ?? DEFAULT_COLUMN_WIDTH))
      if (current !== original) {
        changes.push({
          sheetName: sheet.name,
          colIndex: index,
          widthPx: current
        })
      }
    }
  }
  return changes
}

function markWorkbookSaved(serverData = {}) {
  currentWorkbookVersion.value = serverData?.version ? String(serverData.version) : currentWorkbookVersion.value
  currentFileName.value = serverData?.currentFileName || currentFileName.value
  workbookName.value = serverData?.fileName || currentFileName.value || workbookName.value
  pendingStructures.value = []

  for (const sheet of sheets.value) {
    for (const cell of sheet.cells) {
      cell.originalValue = normalizeValue(cell.value)
      cell.originalStyle = cloneCellStyle(cell.style)
      cell.dirty = false
      cell.forceDirty = false
    }
    sheet.originalMergeKeys = collectSheetMergeKeys(sheet)
    sheet.originalRowHeights = sheet.rowHeights.map((height) => Math.round(Number(height ?? DEFAULT_ROW_HEIGHT)))
    sheet.originalColumnWidths = sheet.columnWidths.map((width) => Math.round(Number(width ?? DEFAULT_COLUMN_WIDTH)))
    initializeSheetDirtyMetrics(sheet)
  }

  refreshDirtyState()
}

async function saveToServer(options = {}) {
  const silent = Boolean(options?.silent)
  const commitLocal = Boolean(options?.commitLocal)
  if (!currentFileName.value || !hasWorkbook.value) {
    if (!silent) {
      ElMessage.warning('璇峰厛鍔犺浇 Excel 鏂囦欢')
    }
    return
  }
  if (!currentWorkbookVersion.value) {
    if (!silent) {
      ElMessage.warning('褰撳墠鏂囦欢鐗堟湰淇℃伅缂哄け锛岃鍏堥噸鏂板姞杞藉悗鍐嶄繚瀛?)
    }
    await loadWorkbookByFileName(currentFileName.value, true, { notify: !silent })
    return
  }

  if (!lockState.value.self || lockState.value.fileName !== currentFileName.value) {
    const lockAcquired = await acquireEditorLock(currentFileName.value)
    if (!lockAcquired) {
      return
    }
  }

  const changes = collectChanges()
  const mergeChanges = countMergeChanges()
  const rowHeightChanges = collectRowHeightChanges()
  const columnWidthChanges = collectColumnWidthChanges()
  const structures = pendingStructures.value.slice()
  if (!changes.length && !mergeChanges && !rowHeightChanges.length && !columnWidthChanges.length && !structures.length) {
    if (!silent) {
      ElMessage.info('褰撳墠娌℃湁闇€瑕佷繚瀛樼殑淇敼')
    }
    return
  }

  saving.value = true
  try {
    const response = await applyExcelWorkbookChanges({
      fileName: currentFileName.value,
      version: currentWorkbookVersion.value,
      changes,
      mergeRegions: mergeChanges ? collectMergeRegions() : undefined,
      rowHeights: rowHeightChanges.length ? rowHeightChanges : undefined,
      columnWidths: columnWidthChanges.length ? columnWidthChanges : undefined,
      structures: structures.length ? structures : undefined
    })
    if (response?.code !== 200) {
      if (response?.data?.versionConflict) {
        const nextFileName = response?.data?.currentFileName || currentFileName.value
        if (!silent) {
          ElMessage.warning(response?.msg || '鏂囦欢宸插湪鏈嶅姟鍣ㄦ洿鏂帮紝璇烽噸鏂板姞杞藉悗鍐嶄繚瀛?)
        }
        await loadWorkbookByFileName(nextFileName, true, { notify: !silent })
        return
      }
      if (!silent) {
        ElMessage.error(response?.msg || '淇濆瓨宸ヤ綔绨垮け璐?)
      }
      return
    }
    if (commitLocal) {
      markWorkbookSaved(response?.data || {})
    } else {
      const nextFileName = response?.data?.currentFileName || currentFileName.value
      await loadWorkbookByFileName(nextFileName, true, { notify: !silent })
    }
    if (!silent) {
      ElMessage.success(response?.msg || 'Excel 淇濆瓨鎴愬姛')
    }
  } catch (error) {
    if (!silent) {
      ElMessage.error(error?.message || '淇濆瓨宸ヤ綔绨垮け璐?)
    }
  } finally {
    saving.value = false
  }
}

async function reloadWorkbook() {
  if (!currentFileName.value) {
    ElMessage.warning('褰撳墠鏈€夋嫨鏂囦欢')
    return
  }
  await loadWorkbookByFileName(currentFileName.value)
}

async function downloadServerWorkbook() {
  if (!currentFileName.value) {
    ElMessage.warning('褰撳墠娌℃湁鍙笅杞芥枃浠?)
    return
  }
  try {
    const buffer = await getExcelEditorContent(currentFileName.value)
    saveAs(new Blob([buffer]), currentFileName.value)
    ElMessage.success('鏂囦欢涓嬭浇鎴愬姛')
  } catch (error) {
    ElMessage.error(error?.message || '涓嬭浇澶辫触')
  }
}

function hasBackgroundImages(sheet) {
  return sheetHasBackgroundImage(sheet)
}

function hasAnchoredImages(sheet) {
  const images = Array.isArray(sheet?.images) ? sheet.images : []
  return images.some((image) => image.kind === 'anchored')
}

function listBackgroundImages(sheet) {
  const images = Array.isArray(sheet?.images) ? sheet.images : []
  return images.filter((image) => image.kind === 'background')
}

function listAnchoredImages(sheet) {
  const images = Array.isArray(sheet?.images) ? sheet.images : []
  return images.filter((image) => image.kind === 'anchored')
}

function sheetHasBackgroundImage(sheet) {
  const images = Array.isArray(sheet?.images) ? sheet.images : []
  return images.some((image) => image.kind === 'background')
}

function buildImageKey(image, index) {
  return `${image.kind}:${image.row1}:${image.col1}:${image.widthPx}:${image.heightPx}:${index}`
}

function buildSheetStageStyle(sheet) {
  return {
    width: `${resolveSheetPixelWidth(sheet)}px`,
    minWidth: '100%',
    minHeight: `${resolveSheetPixelHeight(sheet)}px`
  }
}

function buildImageStyle(image) {
  const sheet = currentSheet.value
  if (!sheet) {
    return {}
  }

  const dx = Math.max(Number(image?.dx1Px ?? 0), 0)
  const dy = Math.max(Number(image?.dy1Px ?? 0), 0)
  const top = COLUMN_HEADER_HEIGHT + resolveFrozenTopOffset(sheet, image.row1) + dy
  const left = ROW_INDEX_WIDTH + resolveFrozenLeftOffset(sheet, image.col1) + dx

  const style = {
    top: `${top}px`,
    left: `${left}px`,
    width: `${Math.max(Number(image?.widthPx ?? 0), 1)}px`,
    height: `${Math.max(Number(image?.heightPx ?? 0), 1)}px`
  }

  const freezePane = sheet.freezePane || { xSplit: 0, ySplit: 0 }
  const freezeRows = Math.max(Number(freezePane.ySplit ?? 0), 0)
  const freezeColumns = Math.max(Number(freezePane.xSplit ?? 0), 0)
  const inFrozenRow = image.kind !== 'background' && image.row1 < freezeRows
  const inFrozenColumn = image.kind !== 'background' && image.col1 < freezeColumns
  if (inFrozenRow || inFrozenColumn) {
    style.position = 'sticky'
    if (inFrozenRow) {
      style.top = `${top}px`
    }
    if (inFrozenColumn) {
      style.left = `${left}px`
    }
    style.zIndex = inFrozenRow && inFrozenColumn ? 8 : 7
  }
  return style
}

function buildCellTdStyle(cell) {
  const sheet = currentSheet.value
  const defaultBackground = sheet && sheetHasBackgroundImage(sheet) ? 'transparent' : '#ffffff'
  const style = {
    width: `${cell.widthPx}px`,
    minWidth: `${cell.widthPx}px`,
    height: `${cell.heightPx}px`,
    backgroundColor: cell.style?.backgroundColor || defaultBackground,
    textAlign: cell.style?.textAlign || 'left',
    verticalAlign: cell.style?.verticalAlign || 'top',
    borderTop: cell.style?.borderTop || '',
    borderRight: cell.style?.borderRight || '',
    borderBottom: cell.style?.borderBottom || '',
    borderLeft: cell.style?.borderLeft || ''
  }

  if (!sheet) {
    return style
  }

  const freezePane = sheet.freezePane || { xSplit: 0, ySplit: 0 }
  const freezeRows = Math.max(Number(freezePane.ySplit ?? 0), 0)
  const freezeColumns = Math.max(Number(freezePane.xSplit ?? 0), 0)
  const inFrozenRow = cell.rowIndex < freezeRows
  const inFrozenColumn = cell.colIndex < freezeColumns
  const atFreezeBottomEdge = freezeRows > 0 && cell.rowIndex + cell.rowSpan - 1 === freezeRows - 1
  const atFreezeRightEdge = freezeColumns > 0 && cell.colIndex + cell.colSpan - 1 === freezeColumns - 1

  if (inFrozenRow || inFrozenColumn) {
    style.position = 'sticky'
    if (inFrozenRow) {
      style.top = `${COLUMN_HEADER_HEIGHT + resolveFrozenTopOffset(sheet, cell.rowIndex)}px`
    }
    if (inFrozenColumn) {
      style.left = `${ROW_INDEX_WIDTH + resolveFrozenLeftOffset(sheet, cell.colIndex)}px`
    }
    style.zIndex = inFrozenRow && inFrozenColumn ? 5 : 4
    style.boxShadow = 'inset 0 0 0 1px rgba(15, 122, 90, 0.08)'
  }
  if (atFreezeBottomEdge && !cell.style?.borderBottom) {
    style.borderBottom = '2px solid #2563eb'
  }
  if (atFreezeRightEdge && !cell.style?.borderRight) {
    style.borderRight = '2px solid #2563eb'
  }

  return style
}

function buildCellClass(cell) {
  return {
    selected: isCellSelected(cell),
    'has-validation': hasValidationOptions(cell),
    'fill-preview': isCellInFillPreview(cell)
  }
}

function buildRowIndexStyle(row) {
  const style = {
    height: `${row.heightPx}px`
  }

  const sheet = currentSheet.value
  if (!sheet) {
    return style
  }
  const freezeRows = Math.max(Number(sheet.freezePane?.ySplit ?? 0), 0)
  if (row.rowIndex < freezeRows) {
    style.top = `${COLUMN_HEADER_HEIGHT + resolveFrozenTopOffset(sheet, row.rowIndex)}px`
    style.zIndex = 7
  }
  if (freezeRows > 0 && row.rowIndex === freezeRows - 1) {
    style.borderBottom = '2px solid #2563eb'
  }
  return style
}

function buildColumnHeaderStyle(sheet, columnIndex) {
  const width = Math.max(Number(sheet?.columnWidths?.[columnIndex] ?? DEFAULT_COLUMN_WIDTH), MIN_COLUMN_WIDTH)
  const style = {
    width: `${width}px`,
    minWidth: `${width}px`,
    height: `${COLUMN_HEADER_HEIGHT}px`
  }
  const freezeColumns = Math.max(Number(sheet?.freezePane?.xSplit ?? 0), 0)
  if (columnIndex < freezeColumns) {
    style.left = `${ROW_INDEX_WIDTH + resolveFrozenLeftOffset(sheet, columnIndex)}px`
    style.zIndex = 10
  }
  if (freezeColumns > 0 && columnIndex === freezeColumns - 1) {
    style.borderRight = '2px solid #2563eb'
  }
  return style
}

function buildCellInputStyle(cell) {
  return {
    minHeight: `${Math.max(cell.heightPx - 8, 28)}px`,
    height: `${Math.max(cell.heightPx - 8, 28)}px`,
    textAlign: cell.style?.textAlign || 'left',
    color: cell.style?.color || '',
    fontWeight: cell.style?.fontWeight || '',
    fontStyle: cell.style?.fontStyle || '',
    fontSize: cell.style?.fontSize || '',
    fontFamily: cell.style?.fontFamily || '',
    textDecoration: cell.style?.textDecoration || '',
    whiteSpace: cell.style?.whiteSpace || 'nowrap'
  }
}

function resolveFrozenTopOffset(sheet, rowIndex) {
  let total = 0
  for (let index = 0; index < rowIndex; index += 1) {
    total += Math.max(Number(sheet.rowHeights[index] ?? DEFAULT_ROW_HEIGHT), MIN_ROW_HEIGHT)
  }
  return total
}

function resolveFrozenLeftOffset(sheet, columnIndex) {
  let total = 0
  for (let index = 0; index < columnIndex; index += 1) {
    total += Math.max(Number(sheet.columnWidths[index] ?? DEFAULT_COLUMN_WIDTH), MIN_COLUMN_WIDTH)
  }
  return total
}

function resolveSheetPixelWidth(sheet) {
  if (!sheet) {
    return ROW_INDEX_WIDTH + DEFAULT_COLUMN_WIDTH
  }
  let total = ROW_INDEX_WIDTH
  const maxColumnCount = Math.max(Number(sheet.maxColumnCount ?? 0), 1)
  for (let index = 0; index < maxColumnCount; index += 1) {
    total += Math.max(Number(sheet.columnWidths[index] ?? DEFAULT_COLUMN_WIDTH), MIN_COLUMN_WIDTH)
  }
  return total
}

function resolveSheetPixelHeight(sheet) {
  if (!sheet) {
    return COLUMN_HEADER_HEIGHT + DEFAULT_ROW_HEIGHT
  }
  let total = COLUMN_HEADER_HEIGHT
  const rowCount = Math.max(Number(sheet.rowCount ?? 0), 1)
  for (let index = 0; index < rowCount; index += 1) {
    total += Math.max(Number(sheet.rowHeights[index] ?? DEFAULT_ROW_HEIGHT), MIN_ROW_HEIGHT)
  }
  return total
}

function hasValidationOptions(cell) {
  return Array.isArray(cell?.validationOptions) && cell.validationOptions.length > 0
}

function buildCellDatalistId(cell) {
  const sheet = activeSheetName.value || 'sheet'
  const safeSheetName = String(sheet).replace(/[^a-zA-Z0-9_-]/g, '_')
  return `excel-list-${safeSheetName}-${cell.rowIndex}-${cell.colIndex}`
}

function startColumnResize(colIndex, event) {
  if (!currentSheet.value) {
    return
  }
  stopFillHandleDrag(false)
  pushUndoSnapshot()
  event.preventDefault()
  event.stopPropagation()
  const sheet = currentSheet.value
  columnResizeState.value = {
    sheetName: sheet.name,
    colIndex,
    startX: Number(event.clientX || 0),
    startWidth: Math.max(Number(sheet.columnWidths[colIndex] ?? DEFAULT_COLUMN_WIDTH), MIN_COLUMN_WIDTH)
  }
  rowResizeState.value = null
  document.body.style.cursor = 'col-resize'
  window.addEventListener('mousemove', handleResizePointerMove)
  window.addEventListener('mouseup', stopResize)
}

function startRowResize(rowIndex, event) {
  if (!currentSheet.value) {
    return
  }
  stopFillHandleDrag(false)
  pushUndoSnapshot()
  event.preventDefault()
  event.stopPropagation()
  const sheet = currentSheet.value
  rowResizeState.value = {
    sheetName: sheet.name,
    rowIndex,
    startY: Number(event.clientY || 0),
    startHeight: Math.max(Number(sheet.rowHeights[rowIndex] ?? DEFAULT_ROW_HEIGHT), MIN_ROW_HEIGHT)
  }
  columnResizeState.value = null
  document.body.style.cursor = 'row-resize'
  window.addEventListener('mousemove', handleResizePointerMove)
  window.addEventListener('mouseup', stopResize)
}

function handleResizePointerMove(event) {
  const sheet = currentSheet.value
  if (!sheet) {
    clearResizeState()
    return
  }

  if (columnResizeState.value && columnResizeState.value.sheetName === sheet.name) {
    const state = columnResizeState.value
    const delta = Number(event.clientX || 0) - state.startX
    const nextWidth = Math.max(MIN_COLUMN_WIDTH, Math.round(state.startWidth + delta))
    if (sheet.columnWidths[state.colIndex] !== nextWidth) {
      sheet.columnWidths[state.colIndex] = nextWidth
      updateCellsForColumnResize(sheet, state.colIndex)
      updateColumnDirtyMarker(sheet, state.colIndex)
      scheduleDirtyStateRefresh()
    }
    return
  }

  if (rowResizeState.value && rowResizeState.value.sheetName === sheet.name) {
    const state = rowResizeState.value
    const delta = Number(event.clientY || 0) - state.startY
    const nextHeight = Math.max(MIN_ROW_HEIGHT, Math.round(state.startHeight + delta))
    if (sheet.rowHeights[state.rowIndex] !== nextHeight) {
      sheet.rowHeights[state.rowIndex] = nextHeight
      if (Array.isArray(sheet.rows) && sheet.rows[state.rowIndex]) {
        sheet.rows[state.rowIndex].heightPx = nextHeight
      }
      updateCellsForRowResize(sheet, state.rowIndex)
      updateRowDirtyMarker(sheet, state.rowIndex)
      scheduleDirtyStateRefresh()
    }
  }
}

function updateCellsForColumnResize(sheet, colIndex) {
  for (const cell of sheet.cells) {
    if (colIndex < cell.colIndex || colIndex >= cell.colIndex + cell.colSpan) {
      continue
    }
    cell.widthPx = resolveRangeWidthFast(sheet, cell.colIndex, cell.colSpan)
  }
}

function updateCellsForRowResize(sheet, rowIndex) {
  for (const cell of sheet.cells) {
    if (rowIndex < cell.rowIndex || rowIndex >= cell.rowIndex + cell.rowSpan) {
      continue
    }
    cell.heightPx = resolveRangeHeightFast(sheet, cell.rowIndex, cell.rowSpan)
  }
}

function resolveRangeWidthFast(sheet, startColumnIndex, colSpan) {
  let total = 0
  for (let offset = 0; offset < colSpan; offset += 1) {
    total += Math.max(Number(sheet.columnWidths[startColumnIndex + offset] ?? DEFAULT_COLUMN_WIDTH), MIN_COLUMN_WIDTH)
  }
  return total
}

function resolveRangeHeightFast(sheet, startRowIndex, rowSpan) {
  let total = 0
  for (let offset = 0; offset < rowSpan; offset += 1) {
    total += Math.max(Number(sheet.rowHeights[startRowIndex + offset] ?? DEFAULT_ROW_HEIGHT), MIN_ROW_HEIGHT)
  }
  return total
}

function stopResize() {
  clearResizeState()
}

function clearResizeState() {
  columnResizeState.value = null
  rowResizeState.value = null
  document.body.style.cursor = ''
  window.removeEventListener('mousemove', handleResizePointerMove)
  window.removeEventListener('mouseup', stopResize)
}

function normalizeValue(value) {
  return value == null ? '' : String(value)
}

function normalizeHexColor(value) {
  const text = normalizeValue(value).trim()
  const matched = text.match(/^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/)
  if (!matched) {
    return ''
  }
  let hex = matched[1].toUpperCase()
  if (hex.length === 3) {
    hex = hex.split('').map((char) => `${char}${char}`).join('')
  }
  return `#${hex}`
}

function normalizeColorStyleValue(value) {
  const hex = normalizeHexColor(value)
  if (hex) {
    return hex
  }
  return convertRgbColorToHex(value)
}

function convertRgbColorToHex(value) {
  const text = normalizeValue(value).trim()
  const matched = text.match(/^rgba?\(([^)]+)\)$/i)
  if (!matched) {
    return ''
  }
  const segments = matched[1]
    .split(',')
    .map((segment) => Number.parseFloat(segment.trim()))
  if (segments.length < 3) {
    return ''
  }
  if (segments.some((segment) => Number.isNaN(segment))) {
    return ''
  }
  const [red, green, blue] = segments
  const toHex = (channel) => {
    const safe = Math.min(255, Math.max(0, Math.round(channel)))
    return safe.toString(16).padStart(2, '0').toUpperCase()
  }
  return `#${toHex(red)}${toHex(green)}${toHex(blue)}`
}

function normalizeEditableStyleValue(key, value) {
  const text = normalizeValue(value).trim()
  if (!text) {
    return ''
  }
  switch (key) {
    case 'textAlign': {
      const lower = text.toLowerCase()
      if (lower === 'left') {
        return ''
      }
      return ['center', 'right', 'justify'].includes(lower) ? lower : ''
    }
    case 'verticalAlign': {
      const lower = text.toLowerCase()
      if (lower === 'top') {
        return ''
      }
      return ['middle', 'bottom'].includes(lower) ? lower : ''
    }
    case 'color':
    case 'backgroundColor':
      return normalizeColorStyleValue(text)
    case 'fontWeight': {
      const lower = text.toLowerCase()
      return lower === '700' || lower === 'bold' ? '700' : ''
    }
    case 'fontStyle':
      return text.toLowerCase() === 'italic' ? 'italic' : ''
    case 'textDecoration':
      return text.toLowerCase().includes('underline') ? 'underline' : ''
    case 'fontSize': {
      const matched = text.match(/^(\d+(?:\.\d+)?)\s*pt$/i)
      if (!matched) {
        return ''
      }
      const sizeValue = Number(matched[1])
      if (!Number.isFinite(sizeValue) || sizeValue < 6 || sizeValue > 72) {
        return ''
      }
      return `${sizeValue}pt`
    }
    case 'fontFamily':
      return text
    case 'whiteSpace':
      return text === 'pre-wrap' ? 'pre-wrap' : ''
    case 'borderTop':
    case 'borderRight':
    case 'borderBottom':
    case 'borderLeft': {
      if (text.toLowerCase() === 'none') {
        return ''
      }
      const matched = text.match(/^(\d+(?:\.\d+)?)px\s+(solid|dashed|dotted)\s+(.+)$/i)
      if (!matched) {
        return ''
      }
      const width = Number(matched[1])
      if (!Number.isFinite(width) || width <= 0 || width > 8) {
        return ''
      }
      const color = normalizeColorStyleValue(matched[3])
      if (!color) {
        return ''
      }
      return `${Math.round(width)}px ${matched[2].toLowerCase()} ${color}`
    }
    case 'numberFormat': {
      const lower = text.toLowerCase()
      if (['percent', 'date', 'number'].includes(lower)) {
        return lower
      }
      return ''
    }
    default:
      return text
  }
}

function sanitizeEditableStyle(style) {
  const source = style && typeof style === 'object' ? style : {}
  const normalized = {}
  for (const key of EDITABLE_STYLE_KEYS) {
    const nextValue = normalizeEditableStyleValue(key, source[key])
    if (nextValue !== '') {
      normalized[key] = nextValue
    }
  }
  return normalized
}

function cloneCellStyle(style) {
  return style && typeof style === 'object' ? { ...style } : {}
}

function buildStylePatchPayload(style) {
  const source = style && typeof style === 'object' ? style : {}
  const payload = {}
  for (const key of EDITABLE_STYLE_KEYS) {
    payload[key] = normalizeEditableStyleValue(key, source[key])
  }
  return payload
}

function hasStyleChanged(cell) {
  const currentStyle = sanitizeEditableStyle(cell?.style)
  const originalStyle = sanitizeEditableStyle(cell?.originalStyle)
  return JSON.stringify(currentStyle) !== JSON.stringify(originalStyle)
}

function applyCellStyleChange(sheet, cell, stylePatch) {
  if (!sheet || !cell || !stylePatch || typeof stylePatch !== 'object') {
    return
  }
  const nextStyle = cloneCellStyle(cell.style)
  for (const [styleKey, styleValue] of Object.entries(stylePatch)) {
    if (!EDITABLE_STYLE_KEYS.includes(styleKey)) {
      continue
    }
    const normalized = normalizeEditableStyleValue(styleKey, styleValue)
    if (normalized === '') {
      delete nextStyle[styleKey]
    } else {
      nextStyle[styleKey] = normalized
    }
  }
  cell.style = nextStyle
  applyCellDirtyChange(sheet, cell, cell.value)
}

function normalizeQueryFileName(value) {
  if (Array.isArray(value)) {
    return value[0] || ''
  }
  return value || ''
}

function getCellRange(cell) {
  return {
    startRow: cell.rowIndex,
    endRow: cell.rowIndex + cell.rowSpan - 1,
    startCol: cell.colIndex,
    endCol: cell.colIndex + cell.colSpan - 1
  }
}

function getSelectionArea(range) {
  return (range.endRow - range.startRow + 1) * (range.endCol - range.startCol + 1)
}

function isMergedCell(cell) {
  return cell.rowSpan > 1 || cell.colSpan > 1
}

function rangesIntersect(left, right) {
  return left.startRow <= right.endRow
    && left.endRow >= right.startRow
    && left.startCol <= right.endCol
    && left.endCol >= right.startCol
}

function formatSelectionRange(range) {
  const startLabel = `${toColumnLabel(range.startCol)}${range.startRow + 1}`
  const endLabel = `${toColumnLabel(range.endCol)}${range.endRow + 1}`
  return startLabel === endLabel ? startLabel : `${startLabel}:${endLabel}`
}

function toColumnLabel(columnIndex) {
  let label = ''
  let current = Math.max(Number(columnIndex ?? 0), 0) + 1
  while (current > 0) {
    const remainder = (current - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    current = Math.floor((current - 1) / 26)
  }
  return label || 'A'
}

function ensureArraySize(target, size, defaultValue) {
  while (target.length < size) {
    target.push(defaultValue)
  }
}

function ensureSheetBounds(sheet, requiredRowCount, requiredColumnCount) {
  sheet.rowCount = Math.max(sheet.rowCount, requiredRowCount, 1)
  sheet.maxColumnCount = Math.max(sheet.maxColumnCount, requiredColumnCount, 1)
  ensureArraySize(sheet.rowHeights, sheet.rowCount, DEFAULT_ROW_HEIGHT)
  ensureArraySize(sheet.columnWidths, sheet.maxColumnCount, DEFAULT_COLUMN_WIDTH)
  if (!sheet.freezePane) {
    sheet.freezePane = normalizeFreezePane(null)
  }
  sheet.freezePane.xSplit = Math.min(Math.max(Number(sheet.freezePane.xSplit ?? 0), 0), sheet.maxColumnCount)
  sheet.freezePane.ySplit = Math.min(Math.max(Number(sheet.freezePane.ySplit ?? 0), 0), sheet.rowCount)
}

function resolveRangeWidth(sheet, startColumnIndex, colSpan) {
  ensureSheetBounds(sheet, sheet.rowCount, startColumnIndex + colSpan)
  let total = 0
  for (let offset = 0; offset < colSpan; offset += 1) {
    total += Math.max(Number(sheet.columnWidths[startColumnIndex + offset] ?? DEFAULT_COLUMN_WIDTH), MIN_COLUMN_WIDTH)
  }
  return total
}

function resolveRangeHeight(sheet, startRowIndex, rowSpan) {
  ensureSheetBounds(sheet, startRowIndex + rowSpan, sheet.maxColumnCount)
  let total = 0
  for (let offset = 0; offset < rowSpan; offset += 1) {
    total += Math.max(Number(sheet.rowHeights[startRowIndex + offset] ?? DEFAULT_ROW_HEIGHT), MIN_ROW_HEIGHT)
  }
  return total
}

function syncCellDimensions(sheet, cell) {
  ensureSheetBounds(sheet, cell.rowIndex + cell.rowSpan, cell.colIndex + cell.colSpan)
  cell.widthPx = resolveRangeWidth(sheet, cell.colIndex, cell.colSpan)
  cell.heightPx = resolveRangeHeight(sheet, cell.rowIndex, cell.rowSpan)
  const validation = resolveCellValidationOptions(sheet, cell.rowIndex, cell.colIndex)
  cell.validationOptions = validation.options
  cell.validationAllowBlank = validation.allowBlank
}

function resolveCellValidationOptions(sheet, rowIndex, colIndex) {
  const validations = Array.isArray(sheet?.validations) ? sheet.validations : []
  for (const validation of validations) {
    if (rowIndex < validation.firstRow || rowIndex > validation.lastRow) {
      continue
    }
    if (colIndex < validation.firstColumn || colIndex > validation.lastColumn) {
      continue
    }
    return {
      options: Array.isArray(validation.options) ? validation.options : [],
      allowBlank: validation.allowBlank !== false
    }
  }
  return {
    options: [],
    allowBlank: true
  }
}

function shouldKeepCurrentValidationValue(cell) {
  if (!cell || !normalizeValue(cell.value)) {
    return false
  }
  return !cell.validationOptions.includes(cell.value)
}

function dedupeCells(cells) {
  const cellMap = new Map()
  for (const cell of cells) {
    cellMap.set(buildCellKey(cell.rowIndex, cell.colIndex), cell)
  }
  return Array.from(cellMap.values()).sort((left, right) => {
    if (left.rowIndex !== right.rowIndex) {
      return left.rowIndex - right.rowIndex
    }
    return left.colIndex - right.colIndex
  })
}

function rebuildSheetRows(sheet) {
  let rowCount = Math.max(Number(sheet.rowCount ?? 0), 1)
  let maxColumnCount = Math.max(Number(sheet.maxColumnCount ?? 0), 1)

  for (const cell of sheet.cells) {
    rowCount = Math.max(rowCount, cell.rowIndex + cell.rowSpan)
    maxColumnCount = Math.max(maxColumnCount, cell.colIndex + cell.colSpan)
  }

  ensureSheetBounds(sheet, rowCount, maxColumnCount)
  sheet.cells = dedupeCells(sheet.cells)

  const rows = Array.from({ length: sheet.rowCount }, (_, rowIndex) => ({
    rowIndex,
    heightPx: Math.max(Number(sheet.rowHeights[rowIndex] ?? DEFAULT_ROW_HEIGHT), MIN_ROW_HEIGHT),
    cells: []
  }))

  for (const cell of sheet.cells) {
    syncCellDimensions(sheet, cell)
    rows[cell.rowIndex].cells.push(cell)
  }

  sheet.rows = rows
}

function buildCellKey(rowIndex, colIndex) {
  return `${rowIndex}:${colIndex}`
}

async function goToFileCenter() {
  await releaseActiveLock(false)
  router.push('/excel-editor/files')
}
</script>

<style lang="scss" scoped>
.excel-editor-page {
  min-height: calc(100vh - 84px);
  padding: 16px;
  background: #f6f8fb;
}

.hero-card,
.summary-card,
.editor-shell {
  border: 1px solid #e4e8ef;
  border-radius: 8px;
  background: #fff;
  box-shadow: none;
}

.hero-card {
  display: flex;
  justify-content: space-between;
  gap: 12px;
  margin-bottom: 12px;
  padding: 12px 14px;
}

.eyebrow {
  margin: 0 0 4px;
  font-size: 11px;
  font-weight: 600;
  letter-spacing: 0.08em;
  text-transform: uppercase;
  color: #6b7280;
}

h1 {
  margin: 0;
  font-size: 20px;
  line-height: 1.3;
  color: #111827;
}

.hero-desc {
  margin: 6px 0 0;
  color: #6b7280;
  line-height: 1.4;
  font-size: 13px;
}

.hero-actions,
.toolbar-left,
.toolbar-right {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
}

.hero-actions {
  align-content: flex-start;
  justify-content: flex-end;
}

.summary-grid {
  display: grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: 10px;
  margin-bottom: 12px;
}

.summary-card {
  display: flex;
  flex-direction: column;
  gap: 4px;
  min-height: 0;
  padding: 10px 12px;
}

.summary-label {
  font-size: 11px;
  font-weight: 600;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: #6b7280;
}

.summary-card strong {
  font-size: 14px;
  line-height: 1.35;
  color: #111827;
}

.summary-card span:last-child {
  color: #6b7280;
  line-height: 1.35;
  font-size: 12px;
}

.editor-shell {
  padding: 12px;
}

.toolbar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  margin-bottom: 14px;
}

.toolbar-right {
  align-items: center;
  justify-content: flex-end;
}

.style-toolbar {
  display: flex;
  align-items: center;
  flex-wrap: wrap;
  gap: 8px;
  margin-bottom: 12px;
  padding: 8px 10px;
  border: 1px solid #e5e7eb;
  border-radius: 6px;
  background: #f9fafb;
}

.formula-toolbar {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-bottom: 12px;
  padding: 8px 10px;
  border: 1px solid #e5e7eb;
  border-radius: 6px;
  background: #ffffff;
}

.formula-label {
  min-width: 72px;
  font-size: 12px;
  font-weight: 600;
  color: #4b5563;
}

.formula-input {
  flex: 1;
  min-width: 260px;
}

.formula-preview {
  font-size: 12px;
  color: #6b7280;
  white-space: nowrap;
}

.style-toolbar-title {
  font-size: 12px;
  font-weight: 600;
  color: #4b5563;
  margin-right: 2px;
}

.style-select {
  width: 90px;
}

.style-color-picker {
  :deep(.el-color-picker__trigger) {
    border-color: #d1d5db;
  }
}

.underline-btn {
  text-decoration: underline;
}

.selection-tip {
  font-size: 12px;
  color: #4b5563;
}

.sheet-tabs {
  margin-bottom: 12px;
}

.sheet-actions {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
  margin-bottom: 12px;
}

.grid-wrap {
  overflow: auto;
  max-height: calc(100vh - 300px);
  min-height: 500px;
  border: 1px solid #e5e7eb;
  border-radius: 6px;
  background: #fff;
}

.sheet-stage {
  position: relative;
  width: max-content;
  min-width: 100%;
}

.sheet-grid {
  position: relative;
  z-index: 3;
  min-width: 100%;
  border-collapse: collapse;
}

.sheet-grid td,
.sheet-grid th {
  border: 1px solid #e5e7eb;
}

.sheet-grid thead th {
  position: sticky;
  top: 0;
}

.corner-header {
  position: sticky;
  top: 0;
  left: 0;
  z-index: 11;
  min-width: 56px;
  width: 56px;
  height: 34px;
  background: #f3f4f6;
}

.column-header {
  position: sticky;
  top: 0;
  z-index: 9;
  padding: 0 18px 0 10px;
  text-align: center;
  font-weight: 700;
  color: #374151;
  background: #f3f4f6;
}

.column-label {
  display: inline-block;
  line-height: 1;
}

.sheet-image-layer {
  position: absolute;
  inset: 0;
  pointer-events: none;
}

.sheet-image-layer-background {
  z-index: 1;
}

.sheet-image-layer-foreground {
  z-index: 6;
}

.sheet-image-item {
  position: absolute;
  overflow: hidden;
}

.sheet-image-item.background {
  opacity: 1;
}

.sheet-image {
  width: 100%;
  height: 100%;
  display: block;
  object-fit: fill;
  user-select: none;
}

.sheet-cell {
  position: relative;
  background: #fff;
  transition: box-shadow 0.18s ease, background-color 0.18s ease;
}

.sheet-cell.selected {
  box-shadow: inset 0 0 0 2px rgba(37, 99, 235, 0.65);
  background: rgba(37, 99, 235, 0.06);
}

.sheet-cell.selected .cell-input {
  background: rgba(37, 99, 235, 0.04);
}

.sheet-cell.fill-preview {
  background: rgba(16, 185, 129, 0.08);
  box-shadow: inset 0 0 0 1px rgba(16, 185, 129, 0.45);
}

.filter-trigger {
  position: absolute;
  top: 3px;
  right: 3px;
  width: 16px;
  height: 16px;
  border: 1px solid #d1d5db;
  border-radius: 4px;
  background: #ffffff;
  color: #6b7280;
  font-size: 11px;
  line-height: 1;
  cursor: pointer;
}

.filter-trigger:hover {
  color: #2563eb;
  border-color: #93c5fd;
  background: #eff6ff;
}

.fill-handle {
  position: absolute;
  right: -4px;
  bottom: -4px;
  width: 8px;
  height: 8px;
  border: 1px solid #ffffff;
  border-radius: 2px;
  background: #2563eb;
  cursor: crosshair;
  z-index: 18;
}

.row-index {
  position: sticky;
  left: 0;
  z-index: 3;
  min-width: 56px;
  width: 56px;
  padding: 0 8px;
  text-align: center;
  font-weight: 700;
  color: #374151;
  background: #f3f4f6;
}

.resize-handle {
  position: absolute;
  user-select: none;
}

.col-resize-handle {
  top: 0;
  right: -4px;
  width: 8px;
  height: 100%;
  cursor: col-resize;
}

.row-resize-handle {
  left: 0;
  bottom: -4px;
  width: 100%;
  height: 8px;
  cursor: row-resize;
}

.cell-input {
  display: block;
  width: 100%;
  min-width: 100%;
  padding: 8px 10px;
  border: none;
  outline: none;
  resize: none;
  background: transparent;
  line-height: 1.45;
  overflow: hidden;
}

.cell-select {
  padding-right: 26px;
}

.sheet-cell.has-validation::after {
  content: "v";
  position: absolute;
  right: 8px;
  bottom: 6px;
  font-size: 10px;
  line-height: 1;
  color: rgba(55, 65, 81, 0.6);
  pointer-events: none;
}

.cell-input.dirty {
  background: rgba(255, 245, 214, 0.85);
}

.cell-input.formula {
  box-shadow: inset 0 0 0 1px rgba(47, 119, 208, 0.14);
}

.sheet-context-menu {
  position: fixed;
  z-index: 3000;
  min-width: 156px;
  border: 1px solid #d1d5db;
  border-radius: 6px;
  background: #fff;
  box-shadow: 0 6px 20px rgba(17, 24, 39, 0.16);
  overflow: hidden;
}

.sheet-context-menu button {
  display: block;
  width: 100%;
  padding: 8px 12px;
  border: none;
  background: transparent;
  text-align: left;
  font-size: 13px;
  color: #1f2937;
  cursor: pointer;
}

.sheet-context-menu button:hover {
  background: #eff6ff;
}

.sort-rule-list {
  display: flex;
  flex-direction: column;
  gap: 8px;
  max-height: 240px;
  overflow: auto;
  padding-right: 2px;
}

.sort-rule-item {
  display: flex;
  align-items: center;
  gap: 8px;
}

.sort-rule-column {
  width: 160px;
}

.sort-rule-direction {
  width: 120px;
}

.filter-value-box {
  width: 100%;
  max-height: 220px;
  overflow: auto;
  border: 1px solid #e5e7eb;
  border-radius: 6px;
  padding: 8px 10px;
}

.find-options {
  display: flex;
  flex-wrap: wrap;
  gap: 12px;
}

.shortcut-grid {
  display: flex;
  flex-direction: column;
  gap: 10px;
}

.shortcut-item {
  display: flex;
  align-items: center;
  gap: 10px;
}

.shortcut-item code {
  min-width: 160px;
  padding: 2px 8px;
  border-radius: 4px;
  background: #f3f4f6;
  color: #111827;
}

@media (max-width: 1200px) {
  .summary-grid {
    grid-template-columns: 1fr;
  }
}

@media (max-width: 960px) {
  .excel-editor-page {
    padding: 16px;
  }

  .hero-card,
  .toolbar,
  .toolbar-right,
  .style-toolbar,
  .formula-toolbar {
    flex-direction: column;
    align-items: flex-start;
  }

  .formula-preview {
    white-space: normal;
  }

  .grid-wrap {
    min-height: 360px;
  }
}
</style>




