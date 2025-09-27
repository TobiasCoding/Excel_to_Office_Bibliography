# Excel to Office Bibliography
This script allow use a XLS/XLSX file to save structurated data and transform to XML file that could be imported like bibliography in Microsoft Word.

## Use

1. Download the exe file from https://github.com/TobiasCoding/Excel_to_Office_Bibliography/raw/refs/heads/main/Excel_to_Office_Bibliography.exe
2. Create and edit a XLSX file like: https://github.com/TobiasCoding/Excel_to_Office_Bibliography/raw/refs/heads/main/bibliography_example.xlsx
3. Open Powershell and execute the executable file with this sample code:
```powershell
.\Excel_to_Office_Bibliography.exe bibliography_example.xlsx bibliography.xml
```
4. Open your Word document.
5. Go to References > Manage Sources.
6. In the Source Manager, click Browse.
7. Select your XML file (e.g., Sources.xml) and confirm. You will see its entries in the Master List.
8. Select the sources you want and click Copy to move them to the Current List. Click OK.
9. Now you can use and edit your bibliography in your Word file!

# To developers:
If you edit the python source code, you can compile it with this PyInstaller command (in Windows):
```powershell
python -m PyInstaller -F -w -n "Excel_to_Office_Bibliography" Excel_to_Office_Bibliography.py Bibliography.xml
```

---
Dev. by tobiasrimoli@protonmail.com
