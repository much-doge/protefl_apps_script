// main.gs
/**
 * Main admin orchestrator for ProTEFL registration workbook.
 * Run this function to set up or refresh everything.
 */
function main() {
    initializeSheets();                 // Creates and populates sheets and templates (setupSheets.gs)
    setupAllDropdowns();                // Adds dropdowns to columns (setupDropdowns.gs)
    protectOriginalScheduleColumn();    // Protects the 'Original Schedule' column (autoCounters.gs)
    applyAllStyling();                  // Applies header and other styling (styling.gs)
    applyAllFormulas();                 // Applies all relevant formulas (applyFormulas.gs)
    // Optionally: syncRescheduleCounts(); // Recalculate reschedule count column (autoCounters.gs)
}