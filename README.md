# D365FO Security Comparison
This project allows for the comparison of D365FO security outputs from different versions.

It utilizes the D365FO_GetSecurityFiles PowerShell script available in this repository.

## Example Usage
The syntax to utilize the tool from command prompt is:
D365FOSecurityComparison.exe <sourceFile> <destFile> <outputType>

<sourceFile> and <destFile> should be the names of the zip files created from the D365FO_GetSecurityFiles PowerShell script.
<outputType> designates the type of file to output, acceptable values are: 'docx' or 'xlsx'

For example, this command would compare the zip files SecurityFiles_10-0-20 and SecurityFiles_10-0-21 and export the results to a Microsoft Word document:
D365FOSecurityComparison.exe SecurityFiles_10-0-20.zip SecurityFiles_10-0-21.zip docx

## License
<a href="http://opensource.org/licenses/MIT">MIT-licensed</a>.
