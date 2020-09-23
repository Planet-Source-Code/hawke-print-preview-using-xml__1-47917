:::An easy way to implement Print Preview:::

1. On Form Load, the ADO data control will fetch records from the database and bind to datagrid to be displayed.

2. The data control then clone a copy of the recordset and set to a local variable

3. When user selects Print Preview, the cloned recordset will be saved as an xml file to the system temp folder, with a unique file name created using CoCreateGuid API.

4. Once the xml file is created, it will be launched in the hidden Web Browser control. When the download of the xml file is completed, it will fire a download complete event, after which the print preview screen will be launched.

5. User can change the page orientation, select print type etc. via the print preview screen without using sophisticated Active X controls.

6. For additional security, the footer of the preview screen is always set to empty.  This is to prevent user from locating the xml file using the printed url.  The xml file is deleted each time the application unloads.


