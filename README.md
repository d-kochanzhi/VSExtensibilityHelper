# VSExtensibilityHelper

Basic useful feature list:

Services:

 * DteService
 * MsgBoxService
 * PaneService
 * ServiceLocator
 
 Core:
 
 * BaseEditorFactory
 * BaseEditorPane
 * BaseWinFormsEditorPane
 * BaseWpfEditorPane
 
#### Making custom VSIX Editor for "*.myExtension" files example:

```c#
 public sealed class WinFormsEditorFactory 
 	: BaseEditorFactory<WinFormsEditorEditorPane>
{
}

public class WinFormsEditorEditorPane 
	: BaseWinFormsEditorPane<WinFormsEditorFactory, Controls.WinForms.WinFormUserControl>
{
	#region Methods

    protected override string GetFileExtension()
    {
    	return ".myExtension";
    }
    protected override void LoadFile(string fileName)
    {}
	protected override void SaveFile(string fileName)
    {}    

    #endregion Methods
}
```

Package registration:

```c#
 [ProvideEditorExtension(typeof(Editors.WinFormsEditorFactory), ".myExtension", 50,
      ProjectGuid = "7D346946-3421-48C0-A98A-20790CB68B3C", NameResourceID = 133,
      TemplateDir = @".\NullPath")]
public sealed class VSPackageCustomEditors : Package
{
 	protected override void Initialize()
	{
        base.Initialize();

            /*
             * Registering editor with WinForms control
             * *.win
             */
            base.RegisterEditorFactory(new Editors.WinFormsEditorFactory()); 
    }
}
```
### Sample project
 \src\VSIXProject_Editor

#### Result
![image](https://github.com/d-kochanzhi/VSExtensibilityHelper/raw/master/resources/2017-10-13_16-48-58.png)

![image](https://github.com/d-kochanzhi/VSExtensibilityHelper/raw/master/resources/2017-10-13_16-49-43.png)

### Based on:

 * [Custom Editors in VSXtra](http://dotneteers.net/blogs/divedeeper/archive/2008/09/01/LearnVSXNowPart30.aspx)
 
## License

This project is licensed under the MIT License