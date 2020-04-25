# Edit a ribbon

The best way to edit a ribbon is, for sure, to use CustomOfficeUIEditor but, for instance, if you don't have the tool installed on your computer but well 7-Zip, you can edit the ribbon with 7-Zip!

The MS Office file format is, in fact, an archive. Just rename the file's extension from `.xlsx` to `.zip` if you want to verify this assertion.

So, you can start the 7-Zip interface and open a .xlsx file where a ribbon has been created.

You'll see a folder called `customUI` and, there, a file called `customUI14.xml` (remember, that was the name of our inserted ribbon). By editing that .xml file within 7-Zip (a text editor will be opened), you can see the content and update it. By saving the file from within the editor, 7-Zip will compress it and add the file to the original MS Office file.

![Edit the ribbon with 7-Zip](./images/7-zip-ribbon.png)
