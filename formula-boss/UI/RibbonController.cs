using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;

using ExcelDna.Integration.CustomUI;

using FormulaBoss.Commands;
using Taglo.Excel.Common;

namespace FormulaBoss.UI;

/// <summary>
///     Provides the Formula Boss ribbon tab in Excel with logo branding and editor launch button.
/// </summary>
[ComVisible(true)]
public class RibbonController : ExcelRibbon
{
    private IRibbonUI? _ribbonUi;

    public override string GetCustomUI(string ribbonId) =>
        @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'
                    onLoad='OnRibbonLoad'>
          <ribbon>
            <tabs>
              <tab id='formulaBossTab' label='Formula Boss'>
                <group id='editorGroup' label='Editor'>
                  <button id='openEditor'
                          label='Open Editor'
                          getImage='GetEditorButtonImage'
                          size='large'
                          onAction='OnOpenEditor'
                          screentip='Open Floating Editor'
                          supertip='Open the Formula Boss floating editor (Ctrl+Shift+`)' />
                </group>
                <group id='settingsGroup' label='Settings'>
                  <button id='openSettings'
                          label='Settings'
                          imageMso='ControlProperties'
                          size='normal'
                          onAction='OnOpenSettings'
                          screentip='Formula Boss Settings'
                          supertip='Configure animation style, indent size, and other preferences' />
                  <button id='aboutButton'
                          label='About'
                          imageMso='Info'
                          size='normal'
                          onAction='OnAbout' />
                </group>
                <group id='updateGroup' label='Updates'>
                  <button id='updateNotification'
                          getLabel='GetUpdateLabel'
                          getVisible='GetUpdateVisible'
                          imageMso='Refresh'
                          size='normal'
                          onAction='OnUpdateClick'
                          screentip='Update Available'
                          supertip='A newer version of Formula Boss is available. Click to download.' />
                </group>
              </tab>
            </tabs>
          </ribbon>
        </customUI>";

    public void OnRibbonLoad(IRibbonUI ribbonUi)
    {
        _ribbonUi = ribbonUi;
        UpdateChecker.UpdateAvailable += OnUpdateAvailable;
    }

    public Bitmap GetEditorButtonImage(IRibbonControl control) => (Bitmap)base.LoadImage("logo32");

    public void OnOpenEditor(IRibbonControl control) => ShowFloatingEditorCommand.ShowFloatingEditor();

    public void OnOpenSettings(IRibbonControl control) => ShowFloatingEditorCommand.ShowSettings();

    public void OnAbout(IRibbonControl control) => ShowFloatingEditorCommand.ShowAbout();

    public string GetUpdateLabel(IRibbonControl control) =>
        UpdateChecker.NewVersionAvailable != null
            ? $"Update: v{UpdateChecker.NewVersionAvailable}"
            : "No Updates";

    public bool GetUpdateVisible(IRibbonControl control) =>
        UpdateChecker.NewVersionAvailable != null;

    public void OnUpdateClick(IRibbonControl control)
    {
        if (UpdateChecker.ReleaseUrl != null)
        {
            Process.Start(new ProcessStartInfo(UpdateChecker.ReleaseUrl) { UseShellExecute = true });
        }
    }

    private void OnUpdateAvailable()
    {
        try
        {
            _ribbonUi?.InvalidateControl("updateNotification");
        }
        catch (Exception ex)
        {
            Logger.Info($"Failed to invalidate ribbon control: {ex.Message}");
        }
    }
}
