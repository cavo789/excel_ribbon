# Assign callbacks

This term means: which code (VBA in this case) should be fired when an event is raised.

When the user clicks on a button of the ribbon which subroutine should be called?

The XML code below assign the `OnButtonClicked` subroutine to the `onAction` event of the button.

```xml
<button
    id="customButton"
    label="Button"
    imageMso="AddFolderToFavorites"
    size="large"
    onAction="OnButtonClicked" />
```

So, if we wish to catch that event, we'll need to add a public subroutine in our Excel file, that subroutine can be placed in any module, should be public and should have `OnButtonClicked` as name.

But... depending on the callbacks (click, change, toggle state, change, ...), the definition of the subroutine is not the same.

For a button, the VBA should be something like:

```vbnet
Public Sub OnButtonClicked(control As IRibbonControl)
    ' YOUR CODE
End Sub
```

For an edit box and the `onChange` event, the callback is different:

```xml
<editBox
    id="customEditBox"
    label="Edit Box"
    onChange="OnEditBoxTextChanged" />
```

```vbnet
Public Sub OnEditBoxTextChanged(control As IRibbonControl, sText As String)
    ' YOUR CODE
End Sub
```

The declaration of callbacks can be found on the official site :
[How can I determine the correct signatures for each callback procedure?](
https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/aa722523(v=office.12)#how-can-i-determine-the-correct-signatures-for-each-callback-procedure). Pay attention to the `Signatures` columns; you need to look for `VBA`.
