## WinFormTheme
---
### Usage
Set form colour theme to use.
```vbnet
SetAppTheme(Themes(Theme.System | Theme.Light | Theme.Dark))
```
Apply the theme to given form.
```vbnet
ThemeForm(form as Form)
```
### Application settings
Name|Type|Scope
---|---|---
appTheme|String|User
appAccent|System.Drawing.Color|User