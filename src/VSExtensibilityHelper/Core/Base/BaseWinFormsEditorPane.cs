/*
 https://github.com/d-kochanzhi/VSExtensibilityHelper
 http://dzsoft.ru/post/VSExtensibilityHelper 
 */

using System.Windows.Forms;
using Microsoft.VisualStudio.Shell.Interop;

namespace VSExtensibilityHelper.Core.Base
{
    public abstract class BaseWinFormsEditorPane<TFactory, TUIControl> :
       BaseEditorPane<TFactory>
         where TFactory : IVsEditorFactory
         where TUIControl : UserControl, new()
    {
        #region Fields

        private TUIControl _UIControl;

        #endregion Fields

        #region Constructors

        public BaseWinFormsEditorPane() : base()
        {
            _UIControl = new TUIControl();
        }

        #endregion Constructors

        #region Properties

        public TUIControl UIControl
        {
            get { return _UIControl; }
        }

        public override System.Windows.Forms.IWin32Window Window
        {
            get
            {
                return _UIControl;
            }
        }

        #endregion Properties
    }
}