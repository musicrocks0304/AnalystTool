# Phase 1 Testing Guide

**Purpose**: Validate all 3 critical bug fixes work correctly
**Est. Time**: 30-45 minutes
**Prerequisites**: Fixed xlam file installed, Excel 2010 or later

---

## Test Setup

### Create Test Workbook

1. Open new Excel workbook
2. Save as "DataCleansing_Tests.xlsx"
3. Create the following sheets:
   - Sheet1: "Normal Range Tests"
   - Sheet2: "Filtered Range Tests"
   - Sheet3: "Large Range Tests"
   - Sheet4: "Calculation Mode Tests"

---

## Test Suite 1: Too_Many_Cells Fix

**Bug Fixed**: Function always returned True, ignoring user "No" response

### Test 1.1: Small Range (No Prompt)
**Expected**: No prompt, operation proceeds

**Steps**:
1. Go to Sheet1
2. Enter "  test  " in A1:A10 (10 cells)
3. Select A1:A10
4. Click Analyst Tools → Trim
5. **VERIFY**: No prompt appears, values are trimmed immediately

✅ Pass | ❌ Fail | Notes: _______________

### Test 1.2: Exactly 100,000 Cells (Prompt Appears)
**Expected**: Prompt appears with formatted count

**Steps**:
1. Go to Sheet3
2. Select A1:D25000 (exactly 100,000 cells)
3. Click Analyst Tools → Trim
4. **VERIFY**: Prompt shows "You have selected 100,000 cells..."
5. Note: Count should be formatted with comma
6. Click "Yes"
7. **VERIFY**: Operation proceeds

✅ Pass | ❌ Fail | Notes: _______________

### Test 1.3: Large Range - Click "Yes" (Proceed)
**Expected**: User confirms, operation proceeds

**Steps**:
1. Go to Sheet3
2. Type "  test  " in A1
3. Copy A1, select A1:A150000 (150,000 cells)
4. Paste
5. Select A1:A150000
6. Click Analyst Tools → Trim
7. **VERIFY**: Prompt shows "You have selected 150,000 cells..."
8. Click "Yes"
9. **VERIFY**: Progress bar appears and operation completes

✅ Pass | ❌ Fail | Notes: _______________

### Test 1.4: Large Range - Click "No" (Abort) ⚠️ CRITICAL
**Expected**: User declines, operation aborts

**Steps**:
1. Go to Sheet3
2. Enter "  test  " with spaces in A1:A150000
3. Select A1:A150000
4. Click Analyst Tools → Trim
5. **VERIFY**: Prompt shows "You have selected 150,000 cells..."
6. Click "No"
7. **VERIFY**:
   - No progress bar appears
   - No processing occurs
   - Data remains unchanged with spaces
8. **BEFORE FIX**: Operation would proceed despite clicking "No"

✅ Pass | ❌ Fail | Notes: _______________

### Test 1.5: Overflow Error (Extreme Range)
**Expected**: Error message about Excel limits

**Steps**:
1. Try to select A1:A5000000 (5 million cells)
   - Or select entire column A (Ctrl+Shift+Down)
2. Click Analyst Tools → Trim
3. **VERIFY**: Error message appears about calculation limits
4. **VERIFY**: Operation does not proceed

✅ Pass | ❌ Fail | Notes: _______________

---

## Test Suite 2: Change_Case Sentence Case Fix

**Bug Fixed**: Sentence case didn't work on filtered ranges

### Test 2.1: Sentence Case on Normal Range
**Expected**: Works correctly (was already working)

**Steps**:
1. Go to Sheet1
2. Enter the following in column A:
   - A1: "hello world"
   - A2: "THIS IS A TEST"
   - A3: "another test row"
3. Select A1:A3
4. Click Analyst Tools → Change Case → Sentence case
5. **VERIFY** results:
   - A1: "Hello world"
   - A2: "This is a test"
   - A3: "Another test row"

✅ Pass | ❌ Fail | Notes: _______________

### Test 2.2: All Case Types on Normal Range
**Expected**: All work correctly

**Steps**:
1. Enter "Hello World" in B1:B4
2. Select B1:B4
3. Test each case type:
   - lowercase → "hello world"
   - UPPERCASE → "HELLO WORLD"
   - Proper Case → "Hello World"
   - Sentence case → "Hello world"
4. **VERIFY**: Each works as expected

✅ Pass | ❌ Fail | Notes: _______________

### Test 2.3: Sentence Case on Filtered Range ⚠️ CRITICAL
**Expected**: Works correctly (was broken before fix)

**Steps**:
1. Go to Sheet2
2. Create test data:
   ```
   A1: hello world
   A2: TEST DATA
   A3: another row
   A4: FILTERED OUT
   A5: last row
   ```
3. Click in A1, apply AutoFilter (Data → Filter)
4. Filter column A to hide row 4:
   - Click dropdown in A1
   - Uncheck "FILTERED OUT"
   - Click OK
5. Verify row 4 is hidden
6. Select A1:A5 (includes hidden row 4)
7. Click Analyst Tools → Change Case → Sentence case
8. **VERIFY** visible rows:
   - A1: "Hello world"
   - A2: "Test data"
   - A3: "Another row"
   - A5: "Last row"
9. Remove filter (click dropdown → Clear Filter)
10. **VERIFY** row 4 is unchanged:
    - A4: "FILTERED OUT" (should be unchanged)
11. **BEFORE FIX**: Sentence case would not be applied to any cells

✅ Pass | ❌ Fail | Notes: _______________

### Test 2.4: All Case Types on Filtered Range
**Expected**: All work correctly

**Steps**:
1. With same filtered data (A4 hidden)
2. Test each case type on A1:A5:
   - lowercase → visible rows become lowercase, A4 unchanged
   - UPPERCASE → visible rows become UPPERCASE, A4 unchanged
   - Proper Case → visible rows become Proper, A4 unchanged
   - Sentence case → visible rows sentence case, A4 unchanged
3. **VERIFY**: Each works, hidden rows unchanged

✅ Pass | ❌ Fail | Notes: _______________

---

## Test Suite 3: Application.Calculation Restoration

**Bug Fixed**: Calculation mode not restored if errors occurred

**Test this for ALL 6 procedures**:
1. Blank_Cells
2. Change_Case
3. Flip_Signs
4. Text_To_Value
5. Trim_Range
6. Value_To_Text

### Test 3.1: Normal Execution (Calc Mode Preserved)

**For EACH procedure**:

**Steps**:
1. Go to Sheet4
2. Set calculation mode to Automatic:
   - Formulas tab → Calculation Options → Automatic
3. Enter:
   - A1: 10
   - A2: 20
   - A3: =SUM(A1:A2)
4. Verify A3 shows 30
5. Select appropriate test data for procedure:
   - Trim: Enter "  test  " in B1:B10
   - Change Case: Enter "HELLO" in B1:B10
   - Flip Signs: Enter "100" in B1:B10
   - Text to Value: Enter '123 in B1:B10 (apostrophe prefix)
   - Value to Text: Enter 123 in B1:B10
   - Blank Cells: Leave B1:B10 empty
6. Run the procedure
7. After completion, change A1 to 50
8. **VERIFY**: A3 immediately updates to 70 (calc mode still Automatic)

✅ Trim_Range | ❌ Fail
✅ Change_Case | ❌ Fail
✅ Flip_Signs | ❌ Fail
✅ Text_To_Value | ❌ Fail
✅ Value_To_Text | ❌ Fail
✅ Blank_Cells | ❌ Fail

### Test 3.2: Error During Execution (Calc Mode Restored) ⚠️ CRITICAL

**For EACH procedure**:

**Steps**:
1. Set calculation mode to Automatic
2. Create formula: A3: =SUM(A1:A2)
3. Create test data in B1:B20
4. **Protect part of range**:
   - Select B11:B20
   - Review tab → Allow Users to Edit Ranges
   - Or: Right-click → Format Cells → Protection → Locked
   - Review tab → Protect Sheet → OK (no password)
5. Select B1:B20 (includes protected cells)
6. Run the procedure
7. **VERIFY**: Error occurs when hitting protected cells
8. Dismiss error message
9. Check calculation mode: Formulas → Calculation Options
10. **VERIFY**: Still shows "Automatic"
11. Change A1 value
12. **VERIFY**: A3 updates immediately
13. **BEFORE FIX**: Mode could be stuck in Manual

✅ Trim_Range | ❌ Fail
✅ Change_Case | ❌ Fail
✅ Flip_Signs | ❌ Fail
✅ Text_To_Value | ❌ Fail
✅ Value_To_Text | ❌ Fail
✅ Blank_Cells | ❌ Fail

### Test 3.3: Start in Manual Mode (Returns to Manual)

**For EACH procedure**:

**Steps**:
1. Set calculation mode to **Manual**:
   - Formulas → Calculation Options → Manual
2. Create formula: A3: =SUM(A1:A2)
3. Change A1 - verify A3 does NOT update automatically
4. Create test data in B1:B20
5. Run the procedure
6. After completion, check mode
7. **VERIFY**: Still shows "Manual"
8. Change A1 again
9. **VERIFY**: A3 still does NOT update automatically
10. Press F9 to calculate
11. **VERIFY**: A3 updates

✅ Trim_Range | ❌ Fail
✅ Change_Case | ❌ Fail
✅ Flip_Signs | ❌ Fail
✅ Text_To_Value | ❌ Fail
✅ Value_To_Text | ❌ Fail
✅ Blank_Cells | ❌ Fail

### Test 3.4: Screen Updating Optimization

**For ANY procedure**:

**Steps**:
1. Create large range of test data (e.g., 10,000 cells)
2. Position Excel window so you can see the cells
3. Run procedure
4. **OBSERVE**:
   - Screen should NOT flicker during processing
   - Changes should appear all at once when complete
5. **BEFORE FIX**: Screen would flicker/update continuously

✅ Pass | ❌ Fail | Notes: _______________

---

## Test Suite 4: Regression Testing

**Purpose**: Ensure existing functionality still works

### Test 4.1: Trim Range
**Steps**:
1. Enter "  test  " (spaces on both sides) in A1:A5
2. Select A1:A5
3. Analyst Tools → Trim
4. **VERIFY**: All spaces removed, cells show "test"

✅ Pass | ❌ Fail | Notes: _______________

### Test 4.2: Text to Value
**Steps**:
1. Enter '123 (with apostrophe) in A1:A5
2. Select A1:A5
3. Analyst Tools → Text To Value
4. **VERIFY**:
   - Green triangle indicator gone
   - Can use in formulas (=SUM works)
   - Format shows General

✅ Pass | ❌ Fail | Notes: _______________

### Test 4.3: Value to Text
**Steps**:
1. Enter 123 (number) in A1:A5
2. Select A1:A5
3. Analyst Tools → Value To Text
4. **VERIFY**:
   - Format shows Text (@)
   - Left-aligned (text alignment)
   - Can concatenate with &

✅ Pass | ❌ Fail | Notes: _______________

### Test 4.4: Flip Signs
**Steps**:
1. Enter mixed values:
   - A1: 100
   - A2: -50
   - A3: 0
   - A4: 75
   - A5: -25
2. Select A1:A5
3. Analyst Tools → Flip Signs
4. **VERIFY**:
   - A1: -100
   - A2: 50
   - A3: 0 (unchanged)
   - A4: -75
   - A5: 25

✅ Pass | ❌ Fail | Notes: _______________

### Test 4.5: Blank Cells - Find Only
**Steps**:
1. Enter data with gaps:
   - A1: "test"
   - A2: (blank)
   - A3: "data"
   - A4: (blank)
   - A5: (blank)
2. Clear "Fill Contents With" box
3. Uncheck "Highlight Contents"
4. Select A1:A5
5. Analyst Tools → Blank Cells → Search & Apply
6. **VERIFY**:
   - Blank cells selected (A2, A4, A5)
   - No highlighting
   - No values filled

✅ Pass | ❌ Fail | Notes: _______________

### Test 4.6: Blank Cells - Highlight & Fill
**Steps**:
1. Same data as 4.5
2. Enter "N/A" in "Fill Contents With"
3. Check "Highlight Contents"
4. Select A1:A5
5. Analyst Tools → Blank Cells → Search & Apply
6. **VERIFY**:
   - A2, A4, A5 highlighted in yellow
   - A2, A4, A5 contain "N/A"
   - A1, A3 unchanged

✅ Pass | ❌ Fail | Notes: _______________

### Test 4.7: Autofit Columns
**Steps**:
1. Enter varying length text:
   - A1: "Short"
   - B1: "This is a very long text string"
   - C1: "X"
2. Make all columns narrow (30 pixels)
3. Select A1:C1
4. Analyst Tools → Autofit Columns
5. **VERIFY**: Each column sized to content

✅ Pass | ❌ Fail | Notes: _______________

---

## Test Suite 5: Progress Bar & Stop Button

### Test 5.1: Progress Bar Appears
**Steps**:
1. Select large range (>100K cells)
2. Click Yes to proceed
3. **VERIFY**: Progress bar appears
4. **VERIFY**: Shows procedure name
5. **VERIFY**: Has Stop button

✅ Pass | ❌ Fail | Notes: _______________

### Test 5.2: Stop Button Works
**Steps**:
1. Create very large range with data (500K+ cells)
2. Run Trim operation
3. While processing, click "Stop" button
4. **VERIFY**:
   - Processing stops
   - Progress bar closes
   - Partial results may be saved
   - No error message

✅ Pass | ❌ Fail | Notes: _______________

---

## Test Results Summary

**Date Tested**: _______________
**Tested By**: _______________
**Excel Version**: _______________
**Add-in Version**: Phase 1 Fixed

### Critical Tests Status

| Test | Status | Notes |
|------|--------|-------|
| Too_Many_Cells Abort | ⬜ Pass / ⬜ Fail | |
| Sentence Case Filtered | ⬜ Pass / ⬜ Fail | |
| Calc Restoration (6 procedures) | ⬜ Pass / ⬜ Fail | |

### Overall Results

- Total Tests: 35+
- Passed: _____ / _____
- Failed: _____ / _____
- Blocked: _____ / _____

### Issues Found

1. _______________________________________________
2. _______________________________________________
3. _______________________________________________

### Recommendation

⬜ **APPROVE** - All critical tests passed, ready for production
⬜ **CONDITIONAL** - Minor issues found, approve with notes
⬜ **REJECT** - Critical issues found, needs fixes

### Sign-Off

**Tester**: _______________  **Date**: _______________
**Reviewer**: _______________  **Date**: _______________
**Approver**: _______________  **Date**: _______________

---

## Troubleshooting

### Issue: "Compile Error" when opening VBA Editor

**Solution**:
1. Check that all variable declarations are present
2. Verify no syntax errors in pasted code
3. Ensure module name is correct: modToolsDataCleansing

### Issue: Tool doesn't appear in ribbon

**Solution**:
1. Close Excel completely
2. Delete Excel temp files: %TEMP%
3. Reopen Excel
4. Verify add-in is loaded: File → Options → Add-ins

### Issue: "Method not found" error

**Solution**:
1. Verify all supporting modules are present (modGlobals, etc.)
2. Check that progress bar module exists
3. Ensure no missing references: Tools → References in VBA Editor

### Issue: Tests fail but unsure why

**Solution**:
1. Enable VBA debugging: Tools → Options → General → Break on All Errors
2. Run test again
3. When error occurs, press Debug
4. Check Locals window to see variable values
5. Document exact error message and location

---

**End of Testing Guide**
