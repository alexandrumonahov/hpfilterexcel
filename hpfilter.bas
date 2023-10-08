Attribute VB_Name = "Module1"
' This version of the Excel HP Filter macro was written by Alexandru Monahov
' It builds upon the original filters and add-on developed by Kurt Annen
' This new version has several improvements in functionality:
' 1) It extends to the one-sided HP filter the ability to process several
'    series at the same time. Previously, this functionality was only
'    available in the two-sided HP filter macro implementation.
' 2) It allows users to process data which is structured both vertically
'    (from top to bottom), as well as horizontally (from left to right),
'    by toggling a newly-implemented 'direction' option.
' 3) The macro workbook can be launched easily in later versions of Office
'    which limit the usage of the original add-on to a single session.

Option Explicit
Option Base 1

Function HPTWO(data As Range, lambda As Double, Optional direction As String)
Attribute HPTWO.VB_Description = "Calculates the two-sided Hodrick-Prescott filter for a given range and lambda value"
Attribute HPTWO.VB_ProcData.VB_Invoke_Func = " \n4"
    Dim nobs As Long, nseries As Long
    Dim i As Long, k As Long
    Dim a() As Double, b() As Double, c() As Double, HPout() As Double
    Dim H1() As Double, H2() As Double, H3() As Double, H4() As Double, H5() As Double
    Dim HH1() As Double, HH2() As Double, HH3() As Double, HH4() As Double, HH5() As Double
    Dim Z() As Double, HB() As Double, HC() As Double

' Select direction of data (horizontal or vertical)
    If direction = "horizontal" Or direction = "horizontal_1" Then

        nseries = data.Rows.Count
        nobs = data.Columns.Count
        ReDim HPout(nobs, nseries)
        For k = 1 To nobs Step 1
            For i = 1 To nseries Step 1
                HPout(k, i) = data(i, k)
            Next i
        Next k
        
    Else

        nseries = data.Columns.Count
        nobs = data.Rows.Count
        ReDim HPout(nobs, nseries)
        For k = 1 To nseries Step 1
            For i = 1 To nobs Step 1
                HPout(i, k) = data(i, k)
            Next i
        Next k
    
    End If
    
    ' If number of observations below 3, return original dataset
    If nobs <= 3 Then
    
        HPTWO = HPout
        
    Else
    
        ' Create pentadiagonal Matrix
        ReDim a(nobs, nseries)
        ReDim b(nobs, nseries)
        ReDim c(nobs, nseries)
            
        For k = 1 To nseries Step 1
            a(1, k) = 1 + lambda
            b(1, k) = -2 * lambda
            c(1, k) = lambda
        
            For i = 2 To nobs - 1 Step 1
                a(i, k) = 6 * lambda + 1
                b(i, k) = -4 * lambda
                c(i, k) = lambda
            Next i
        
            a(2, k) = 5 * lambda + 1
            a(nobs, k) = 1 + lambda
            a(nobs - 1, k) = 5 * lambda + 1
            b(1, k) = -2 * lambda
            b(nobs - 1, k) = -2 * lambda
            b(nobs, k) = 0
            c(nobs - 1, k) = 0
            c(nobs, k) = 0
        Next k
        
        ' Solve system of linear equations
        ReDim H1(nseries)
        ReDim H2(nseries)
        ReDim H3(nseries)
        ReDim H4(nseries)
        ReDim H5(nseries)
        ReDim HH1(nseries)
        ReDim HH2(nseries)
        ReDim HH3(nseries)
        ReDim HH4(nseries)
        ReDim HH5(nseries)
        ReDim Z(nseries)
        ReDim HB(nseries)
        ReDim HC(nseries)
        
        For k = 1 To nseries Step 1
        
        ' Forward
            For i = 1 To nobs Step 1
                Z(k) = a(i, k) - H4(k) * H1(k) - HH5(k) * HH2(k)
                HB(k) = b(i, k)
                HH1(k) = H1(k)
                H1(k) = (HB(k) - H4(k) * H2(k)) / Z(k)
                b(i, k) = H1(k)
                HC(k) = c(i, k)
                HH2(k) = H2(k)
                H2(k) = HC(k) / Z(k)
                c(i, k) = H2(k)
                a(i, k) = (HPout(i, k) - HH3(k) * HH5(k) - H3(k) * H4(k)) / Z(k)
                HH3(k) = H3(k)
                H3(k) = a(i, k)
                H4(k) = HB(k) - H5(k) * HH1(k)
                HH5(k) = H5(k)
                H5(k) = HC(k)
            Next i
            
        H2(k) = 0
        H1(k) = a(nobs, k)
        HPout(nobs, k) = H1(k)
        
        ' Backward
            For i = nobs To 1 Step -1
                HPout(i, k) = a(i, k) - b(i, k) * H1(k) - c(i, k) * H2(k)
                H2(k) = H1(k)
                H1(k) = HPout(i, k)
            Next i
            
        Next k
        
        ' Match output direction with input direction
        If direction = "horizontal" Then
            HPTWO = TransposeArray(HPout)
        Else
            HPTWO = HPout
        End If
    
    End If ' End of: If nobs <= 3 Then
    
End Function

Function HPONE(data As Range, lambda As Double, Optional direction As String)
Attribute HPONE.VB_Description = "Calculates the one-sided Hodrick-Prescott filter for a given range and lambda value"
Attribute HPONE.VB_ProcData.VB_Invoke_Func = " \n4"
    Dim nobs As Long, nseries As Long, i As Long, k As Long, j As Long, tmp1 As Variant, tmp2 As Variant, rng As Range
    
    ' Select direction of data (horizontal or vertical)
    If direction = "horizontal" Then
        nobs = data.Rows.Count
        nseries = data.Columns.Count
        ReDim tmp1(1 To nseries, 1 To nobs)
        ' Loop through time series
        For k = 1 To nseries
            For i = 1 To nobs
                Set rng = data.Resize(i, k)
                tmp2 = HPTWO(rng, lambda, "horizontal_1")
                tmp1(k, i) = tmp2(k, i)
            Next i
        Next k
        Set rng = Nothing
        HPONE = TransposeArray(tmp1)
    
    Else
        nobs = data.Rows.Count
        nseries = data.Columns.Count
        ReDim tmp1(1 To nobs, 1 To nseries)
        ' Loop through time series
        For k = 1 To nseries
            For i = 1 To nobs
                Set rng = data.Resize(i, k)
                tmp2 = HPTWO(rng, lambda)
                tmp1(i, k) = tmp2(i, k)
            Next i
        Next k
        Set rng = Nothing
        HPONE = tmp1
        
    End If
    
End Function
    

Function TransposeArray(InputArray As Variant) As Variant
    Dim X As Long, y As Long
    Dim maxX As Long, minX As Long
    Dim maxY As Long, minY As Long
    Dim TempArray As Variant
    
    'Get upper and lower bounds from input array
    maxX = UBound(InputArray, 1)
    minX = LBound(InputArray, 1)
    maxY = UBound(InputArray, 2)
    minY = LBound(InputArray, 2)
    
    'Create a new temporary array
    ReDim TempArray(minY To maxY, minX To maxX)
    
    'Transpose the input array
    For X = minX To maxX
        For y = minY To maxY
            TempArray(y, X) = InputArray(X, y)
        Next y
    Next X
    
    'Output the transposed array
    TransposeArray = TempArray
    
End Function

Private Sub RegisterHPTWO()
    Application.MacroOptions _
        Macro:="HPTWO", _
        Description:="Calculates the two-sided Hodrick-Prescott filter for a given range and lambda value", _
        Category:="Statistical", _
        ArgumentDescriptions:=Array( _
            "Range.  Select the data range, which can include multiple series", _
            "Lambda.  Set the smoothing parameter lambda", _
            "Direction.  (Optional) if variables are in rows and data is in columns from right to left, set to: horizontal")
End Sub

Private Sub RegisterHPONE()
    Application.MacroOptions _
        Macro:="HPONE", _
        Description:="Calculates the one-sided Hodrick-Prescott filter for a given range and lambda value", _
        Category:="Statistical", _
        ArgumentDescriptions:=Array( _
            "Range.  Select the data range, which can include multiple series", _
            "Lambda.  Set the smoothing parameter lambda", _
            "Direction.  (Optional) if variables are in rows and data is in columns from right to left, set to: horizontal")
End Sub




