---
layout: default
title: Excel VBA
nav_order: 3
---
# The challenge
Our team was working with an Excel spreadsheet that was used across several areas. This tool was meant for supervisors and team leaders to manually input problem descriptions and action items. However, this manual process often led to alterations such as adding columns, changing formats, and even breaking the structural integrity of the document.

Here's a look at the basic Excel spreadsheet we started with:

![Basic view of the Excel Spreadsheet](../../assets/images/excel_vba_sheet.png)




# The solution
I stumbled upon Excel VBA (Visual Basic for Applications), a powerful feature that can enhance the functionality of Excel spreadsheets. I decided to learn about it and designed a VBA form where users could simply select from dropdown lists to input what they needed.

This is how the input form looks:

![Input Form](./assets/images/excel_vba_form.png)

I also prioritized data quality by implementing checks and standards at the point of data entry.

To ensure the data quality:

- I pre-loaded the current date into the form.

```vba
Dim TodaysDate As String
   
    TodaysDate = Format(Now(), "dd/mm/yyyy")
   
    DetectionDate.Value = TodaysDate
    ClosureDate.Value = TodaysDate

End Sub
```

- Made sure the estimated resolution date for an entry could not be set earlier than the report creation date.

```vba
Private Sub ClosureDate_afterupdate()
If ClosureDate.Value < DetectionDate.Value Then

 MsgBox "The closure date cannot be earlier than the detection date."
 End If
End Sub
```

![Form Dates](./assets/images/excel_vba_dates.png)

- I locked all sheets, excluding specific columns, and created a simple prompt for supervisors to fix the sheet if any issue arose. Upon closing the document, all pages would automatically be protected again.

```vba
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.DisplayAlerts = False

    'Protect worksheets with passwords
    Sheets("sheet1").Protect password:="password1"
    Sheets("sheet2").Protect password:="password2"
    Sheets("sheet3").Protect password:="password3"
    Sheets("sheet4").Protect password:="password4"
    Sheets("sheet5").Protect password:="password5"
    Sheets("sheet6").Protect password:="password6"
    Sheets("sheet7").Protect password:="password7"
    Sheets("sheet8").Protect password:="password8"
    Sheets("sheet9").Protect password:="password9"

    'Unprotect worksheet for editing
    Sheets("sheet1").Unprotect password:="password1"

    'Lock specific ranges in the worksheet
    Sheets("sheet1").Range("1:6").Locked = True
    Sheets("sheet1").Range("7:9999").Locked = True

    'Re-protect the worksheet
    Sheets("sheet1").Protect password:="password1"

    'Save the workbook
    'ThisWorkbook.Save
End Sub
```

- I also introduced dropdown lists for necessary categories to standardize the inputs.

```vba
Dim index As Integer
index = variable_1.ListIndex

category.Clear

Select Case index
    Case Is = 0
        With category
            .AddItem "var1_cat_1"
            .AddItem "var1_cat_2"
            .AddItem "var1_cat_3"
            .AddItem "var1_cat_4"
            .AddItem "var1_cat_5"
        End With
    ' Add similar cases for different indices
End Select
End Sub
```
{: .highlight }
Data Quality is best done when you apply Data Standards in the origin / input of the information. 

To enhance the user experience:

- I added a validation prompt before pressing the "Cancel" button to prevent accidental loss of information.
Here's how the form looks when the "Cancel" button is clicked:

```vba
Private Sub CANCEL_Click()
result = MsgBox("Cancel input loading? Unsaved data will be lost.", vbYesNo, "Cancel?")
If result = vbYes Then
Unload Me
End If
If result = vbNo Then
Cancel = True
End If
End Sub
```
![Form Cancel](./assets/images/excel_vba_cancel.png)

I then linked these inputs to the visual stats that were previously calculated manually.
Here's a visual representation of the inputs:

![Pyramid](./assets/images/excel_vba_pyramid.png)


```
=+COUNTIFS(PGD[CAT1];VALIDACION!$F$9;PGD[AREA];VALIDACION!$B$3;PGD[DATE];">=" &VLOOKUP(S11;MESES[[#All];[MES]:[FIN]];2;FALSE);PGD[DATE];"<=" &VLOOKUP(S11;MESES[[#All];[MES]:[FIN]];3;FALSE))
```

# The outcome
Thanks to the changes, we experienced significant improvements in the following areas:

- Consistency in input and document formats.
- Increased scope for analysis as a result of unified and normalized categories.
- Enhanced user experience, leading to an increase in action items as the process was more straightforward than before.
- Tracking of open incidents became simpler and more effective.
- Here's a look at the tracking system post-improvement:

![Tracking](./assets/images/excel_vba_open_actions.png)
