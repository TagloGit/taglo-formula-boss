using System.Drawing;
using System.Runtime.InteropServices;

using ExcelDna.Integration.CustomUI;

using FormulaBoss.Commands;

namespace FormulaBoss.UI;

/// <summary>
///     Provides the Formula Boss ribbon tab in Excel with logo branding and editor launch button.
/// </summary>
[ComVisible(true)]
public class RibbonController : ExcelRibbon
{
    public override string GetCustomUI(string ribbonId) =>
        @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
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
              </tab>
            </tabs>
          </ribbon>
        </customUI>";

    public Bitmap GetEditorButtonImage(IRibbonControl control) => (Bitmap)base.LoadImage("logo32");

    public void OnOpenEditor(IRibbonControl control) => ShowFloatingEditorCommand.ShowFloatingEditor();
}
