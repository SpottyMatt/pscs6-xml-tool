# Photoshop CS6 XML Update Tool

A PowerShell script to update the BridgeTalkVersion value (or other numbers) in Adobe Photoshop CS6 application XML files. This tool temporarily disables network connectivity during the update.

## QuickStart

1. Open a command prompt (CMD) or PowerShell terminal
2. Run the batch file:
   ```
   update_xml.bat
   ```
   
   Or specify a custom XML path:
   ```
   update_xml.bat "C:\Path\To\Adobe\Adobe Photoshop CS6 (64 Bit)\AMT\application.xml"
   ```

3. The script will automatically request administrator privileges if needed
4. Network adapters will be temporarily disabled during the update
5. After the XML is updated, press Enter to restore network connectivity
6. Check the generated log file for details of the operation

## User Guide

See included [reference.xml](./reference.xml) for an example of the intended target file.

### Parameters

The script supports several configuration parameters:

| Parameter | Description | Default Value |
|-----------|-------------|---------------|
| XmlFilePath | Path to the Adobe application.xml file | Program Files\Adobe\Adobe Photoshop CS6 (64 Bit)\AMT\application.xml |
| LogFile | Path to the log file | script_folder\update_log.txt |
| DataKey | The XML data key attribute to locate | BridgeTalkVersion |
| AdobeCode | The adobeCode attribute value to search for | /adobe/bridgetalk/photoshop-60.064 |

### Advanced Usage

The batch file supports all the same parameters as the PowerShell script. Here's an example with all parameters:

```batch
update_xml.bat -XmlFilePath "C:\Path\To\application.xml" -DataKey "CustomKey" -AdobeCode "CustomCode" -LogFile "C:\Path\To\custom_log.txt"
```

You can specify any combination of these parameters. If no parameters are provided, the script will use the default values shown in the table above.
