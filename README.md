# MS-MSDT-Office-RCE-Follina
CVE-2022-30190 | MS-MSDT Follina One Click

1. Create a Docx file.
2. In the Docx file, Insert > Object > Bitmap Image > Ok
3. In the Paint application that launched, save the Paint file.
4. Save your Docx file.
5. Open your file as an archive (With 7Zip; Right Click > 7Zip > Open Archive.)
6. Copy out the Document.xml from \Word\ and document.xml.rels file from \word_rels\
7. Open/Edit the \word\_rels\Document.xml.rels and locate the relationship XML Tag with "/relationships/oleObject". After this replace the destination in "Target=" to your remote destination where it will grab the payload from. Also add "TargetMode="External"" to the XML Tag.
