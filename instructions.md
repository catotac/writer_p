# Step-by-Step Guide: Implementing a Questionnaire Ribbon in Word

## Prerequisites
- Microsoft Word (any version that supports VBA, such as Word 2010 or newer)
- Basic familiarity with the Word interface

## Step 1: Open the VBA Editor
1. Open Microsoft Word
2. Press `Alt + F11` on your keyboard
   - This will open the Visual Basic for Applications (VBA) editor

## Step 2: Add a Standard Module
1. In the VBA editor, click on `Insert` in the top menu
2. Select `Module` from the dropdown menu
3. A new module will appear in the Project Explorer (usually on the left side)
4. This will be named "Module1" by default

## Step 3: Add a Class Module
1. In the VBA editor, click on `Insert` in the top menu
2. Select `Class Module` from the dropdown menu
3. Rename this module to "Ribbon":
   - Right-click on the newly created class module in the Project Explorer
   - Select "Properties"
   - Change the name from "Class1" to "Ribbon"
4. Add the following code to this class module:
```vba
Option Explicit

Public ribbon As IRibbonUI

Sub OnLoad(ribbonUI As IRibbonUI)
    Set ribbon = ribbonUI
End Sub
```

## Step 4: Create a User Form
1. In the VBA editor, click on `Insert` in the top menu
2. Select `UserForm` from the dropdown menu
3. Rename this form to "QuestionnaireForm":
   - Right-click on the newly created form in the Project Explorer
   - Select "Properties"
   - Change the name from "UserForm1" to "QuestionnaireForm"

## Step 5: Add the Main Code
1. Double-click on "Module1" in the Project Explorer to open it
2. Copy and paste the main code from the artifact (the part starting with "' This code creates a custom ribbon..." up to but not including the line "' The following goes in the Ribbon class module")

## Step 6: Add Code to the UserForm
1. With the QuestionnaireForm open in design view, right-click on the form and select "View Code"
2. Copy and paste the UserForm code from the artifact (the part starting with "' The following code goes in the code section of the QuestionnaireForm UserForm" and everything after it)

## Step 7: Create the XML File for the Ribbon
1. Return to your Word document
2. Press `Alt + F8` to open the Macros dialog
3. Select "CreateRibbonXML" and click "Run"
4. A message box will appear with instructions for adding the ribbon to Word
5. Take note of the location of the XML file (usually in your Windows Temp folder)

## Step 8: Save Your Document as a Macro-Enabled Template
1. Click on File > Save As
2. Select "Word Macro-Enabled Document (*.docm)" from the file type dropdown
3. Give your document a name and click Save

## Step 9: Add the Custom UI to the Document
1. Close Word completely
2. Navigate to where you saved your .docm file
3. Change the file extension from .docm to .zip (you might need to enable file extensions in File Explorer first)
4. Open the zip file using any archive program (Windows Explorer, 7-Zip, WinRAR, etc.)
5. Create a new folder inside the zip named "customUI" if it doesn't already exist
6. Copy the XML file created in Step 7 into this folder
7. Rename the XML file to "customUI.xml" if it's not already named that

## Step 10: Update the Content Types
1. While still in the zip file, locate the "[Content_Types].xml" file in the root of the archive
2. Open it with a text editor (like Notepad)
3. Before the closing tag `</Types>`, add the following line:
```xml
<Override PartName="/customUI/customUI.xml" ContentType="application/xml"/>
```
4. Save the file and close the text editor

## Step 11: Update the Relationships
1. In the zip file, navigate to the "_rels" folder
2. Open the ".rels" file with a text editor
3. Before the closing tag `</Relationships>`, add the following line:
```xml
<Relationship Id="rId1000" Type="http://schemas.microsoft.com/office/2007/relationships/ui/customUI" Target="customUI/customUI.xml"/>
```
4. Save the file and close the text editor

## Step 12: Finalize the Document
1. Close the zip file
2. Change the file extension back from .zip to .docm
3. Open the document in Word

## Step 13: Enable Macros
1. When opening the document, you may see a security warning about macros
2. Click "Enable Content" to allow the macros to run

## Step 14: Test the Questionnaire
1. You should now see a new tab called "Questionnaire" in the Word ribbon
2. Click on the "Open Questionnaire" button to launch the form
3. Test the form by answering questions and navigating between pages
4. Click "Submit" when done to save your responses to the document

## Troubleshooting
- If the "Questionnaire" tab doesn't appear, check that macros are enabled in Word
- If you encounter errors, verify that all code was copied correctly
- Make sure the XML file was properly added to the customUI folder and the content types were updated correctly

## Notes
- This implementation will only work in the specific document you created
- To make it available for all documents, you would need to create an add-in, which is beyond the scope of these instructions
