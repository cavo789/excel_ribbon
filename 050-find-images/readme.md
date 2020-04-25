# Find images

By adding a button f.i., you'll assign an image to it.

You can use a personal image like the logo of BOSA (in that case, you'll need to import the image in the CustomOfficeUIEditor application and use the `image` attribute).

You'll most probably use an existing, standard, image. This can be done by using the `imageMso` attribute and telling Office which image to use; f.i.  `AddFolderToFavorites`:

![Favorites](./images/Favorites.png)

The manifest is this one:

```xml
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
    <ribbon>
        <tabs>
            <tab idMso="TabHome" visible="false" />
             <tab id="customTab" insertAfterMso="TabView" label="Tab">
                <group id="customGroup" label="Group">
                    <button
                        id="customButton"
                        label="Button"
                        imageMso="AddFolderToFavorites"
                        size="large"
                        onAction="OnButtonClicked" />
                </group>
            </tab>
        </tabs>
    </ribbon>
</customUI>
```

But ... **how to retrieve the list of images?**

Microsoft maintains Excel files with the list of existing IDs that can be used as icons in our ribbon. The "Office 2010 Help Files: Office Fluent User Interface Control Identifiers" can be downloaded [here](https://www.microsoft.com/en-us/download/confirmation.aspx?id=6627). You'll get a lot of Excel files, on file by application (Access, Excel, Outlook, ...).

This will give the list of existing IDs in plain text but you'll not see the associated images.

You can download all images (as PNG) for MS Office 2010 or 2013 [here](http://hintdesk.com/2011/07/22/c-print-all-ms-office-imagemso-to-files) (see links in Chapter 2 Download).

An offline version for Office 2010 is [here](./files/Office_2010_Icon_Gallery_Files.zip) and [here](./files/Office_2013_Icon_Gallery_Files.zip) for Office 2013.
