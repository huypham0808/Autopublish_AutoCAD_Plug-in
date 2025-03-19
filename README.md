## Autopublish_AutoCAD_Plug-in
1. What is Autopublish_AutoCAD_Plug-in?
2. Why should you use this plug-in?
3. How to use this tool?

### 1. What is Autopublish_AutoCAD_Plug-in?
+ Autopublish_AutoCAD_Plug-in is a plug-in for AutoCAD that helps users can export all layouts in a Dwg file into a PDF file by only one COMMAND.
### 2. Why should you use this plug-in?
+ The user does not need to set up the destination file path for the PDF file. The plug-in will get a file path to save a PDF file, the same as the original DWG file.
+ The plug-in will automatically get the current Plot style and apply it to all of the layout sheets.
+ After execution, the plug-in will automatically open the PDF file with the default app to view the result.
### 3. How to setup?
* Clone the project to your local drive and build solution.
* After build the solution, navigate to folder `.\bin\debug` you should see the file  `PCC.dll`
* Open AutoCAD software and command NETLOAD and navigate to the folder storing `.\bin\debug` to load `PCC.dll` file.
  <br>
* If there is any warning Dialog shown as figure below. Please select Always Load button for the first time.
 ![WarningLoad](https://github.com/huypham0808/SD_PrintTool_AutoCAD_Plug-in/assets/114324328/ef3e4182-7502-4c0f-8030-03e98df147e9)
* Now, you guys can enjoy the plug-in.
  <br>
  --- Good luck to my friend---
