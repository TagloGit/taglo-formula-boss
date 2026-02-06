using System.Diagnostics;
using System.Runtime.InteropServices;

using ExcelDna.Integration;

using FormulaBoss.Interception;

namespace FormulaBoss.Commands;

/// <summary>
///     Command handler for editing processed Formula Boss LET formulas.
///     Triggered by Ctrl+Shift+` keyboard shortcut.
/// </summary>
public static class EditFormulaCommand
{
    private const int VkNumlock = 0x90;
    private const int KeyeventfExtendedkey = 0x0001;
    private const int KeyeventfKeyup = 0x0002;

    [DllImport("user32.dll")]
    private static extern short GetKeyState(int keyCode);

    [DllImport("user32.dll")]
    private static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);

    private static bool IsNumLockOn() => (GetKeyState(VkNumlock) & 1) == 1;

    private static void ToggleNumLock()
    {
        keybd_event(VkNumlock, 0x45, KeyeventfExtendedkey, 0);
        keybd_event(VkNumlock, 0x45, KeyeventfExtendedkey | KeyeventfKeyup, 0);
    }

    /// <summary>
    ///     Executes the edit formula command on the active cell.
    ///     If the cell contains a processed Formula Boss LET formula, reconstructs
    ///     the editable version with backtick expressions and enters edit mode.
    /// </summary>
    [ExcelCommand(MenuName = "Formula Boss", MenuText = "Edit Formula")]
    public static void EditFormulaBossFormula()
    {
        try
        {
            dynamic app = ExcelDnaUtil.Application;
            var cell = app.ActiveCell;

            // Get the cell's formula (prefer Formula2 for dynamic arrays)
            var formula = cell.Formula2 as string ?? cell.Formula as string;

            if (string.IsNullOrEmpty(formula))
            {
                Debug.WriteLine("EditFormulaBossFormula: No formula in active cell");
                return;
            }

            Debug.WriteLine($"EditFormulaBossFormula: Processing formula: {formula}");

            if (LetFormulaReconstructor.TryReconstruct(formula, out var editableFormula))
            {
                Debug.WriteLine($"EditFormulaBossFormula: Reconstructed to: {editableFormula}");

                // Temporarily disable events to prevent SheetChange from firing
                // and immediately reprocessing the backtick formula
                app.EnableEvents = false;
                try
                {
                    // Set the cell value as text (the quote prefix makes it text)
                    cell.Value = editableFormula;

                    // Prevent text wrapping so the formula remains readable
                    cell.WrapText = false;
                }
                finally
                {
                    app.EnableEvents = true;
                }

                // Enter edit mode on the cell so user can immediately start editing
                // SendKeys has a known bug that can toggle Num Lock, so we save and restore it
                var numLockWasOn = IsNumLockOn();
                app.SendKeys("{F2}");
                if (IsNumLockOn() != numLockWasOn)
                {
                    ToggleNumLock();
                }
            }
            else
            {
                Debug.WriteLine("EditFormulaBossFormula: Cell does not contain a Formula Boss LET formula");
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"EditFormulaBossFormula error: {ex.Message}");
        }
    }
}
