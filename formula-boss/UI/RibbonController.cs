using System.Drawing;
using System.Reflection;
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
        @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' loadImage='LoadImage'>
          <ribbon>
            <tabs>
              <tab id='formulaBossTab' label='Formula Boss'>
                <group id='editorGroup' label='Editor'>
                  <button id='openEditor'
                          label='Open Editor'
                          image='logo32'
                          size='large'
                          onAction='OnOpenEditor'
                          screentip='Open Floating Editor'
                          supertip='Open the Formula Boss floating editor (Ctrl+Shift+`)' />
                </group>
              </tab>
            </tabs>
          </ribbon>
        </customUI>";

    public Bitmap? LoadImage(string imageId)
    {
        if (imageId == "logo32")
        {
            var stream = Assembly.GetExecutingAssembly()
                .GetManifestResourceStream("FormulaBoss.Resources.logo-32.png");
            return stream != null ? new Bitmap(stream) : null;
        }

        return null;
    }

    public void OnOpenEditor(IRibbonControl control) => ShowFloatingEditorCommand.ShowFloatingEditor();
}
