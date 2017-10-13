/*
 https://github.com/d-kochanzhi/VSExtensibilityHelper
 http://dzsoft.ru/post/VSExtensibilityHelper 
 */
using System.Windows.Controls;
using Microsoft.VisualStudio.Shell.Interop;

namespace VSExtensibilityHelper.Core.Base
{
    public abstract class BaseWpfEditorPane<TFactory, TUIControl> :
      BaseEditorPane<TFactory>
        where TFactory : IVsEditorFactory
        where TUIControl : UserControl, new()
    {
        #region Fields

        private TUIControl _UIControl;

        #endregion Fields

        #region Constructors

        public BaseWpfEditorPane() : base()
        {
            base.Content = _UIControl = new TUIControl();
        }

        #endregion Constructors

        #region Properties

        public TUIControl UIControl
        {
            get { return _UIControl; }
        }

        #endregion Properties
    }
}