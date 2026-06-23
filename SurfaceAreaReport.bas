Option Explicit

' Module-level settings
Dim gFrontPlane As Integer
' *** REMOVED: gRevision, gAuthor, gCompany ***
' *** All metadata (Company, Revision, Author) is now entered directly in the HTML report ***

' =====================================================
' MAIN
' =====================================================
Sub Main()
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then
        MsgBox "Please open a Part or Assembly first!", vbExclamation
        Exit Sub
    End If
    
    Dim planeInput As String
    planeInput = InputBox("Which reference plane is your FRONT?" & vbCrLf & vbCrLf & _
                           "1 = Front Plane (Z axis)" & vbCrLf & _
                           "2 = Top Plane (Y axis)" & vbCrLf & _
                           "3 = Right Plane (X axis)" & vbCrLf & vbCrLf & _
                           "Enter 1, 2, or 3:", "Select Front Plane", "1")
    If planeInput = "" Then Exit Sub
    gFrontPlane = 1
    If IsNumeric(planeInput) Then
        If CLng(planeInput) >= 1 And CLng(planeInput) <= 3 Then gFrontPlane = CInt(planeInput)
    End If
    
    swModel.EditRebuild3
    
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = swModel.SelectionManager
    Dim faceCount As Long, compCount As Long
    faceCount = CountSelectedFaces(swSelMgr)
    compCount = CountSelectedComponents(swSelMgr)
    
    If faceCount > 0 Then
        GenerateFacesReport swModel, swSelMgr, faceCount
        Exit Sub
    End If
    If compCount > 0 And swModel.GetType = swDocASSEMBLY Then
        GenerateSelectedComponentsReport swModel, swSelMgr
        Exit Sub
    End If
    If swModel.GetType = swDocASSEMBLY Then
        Dim visCount As Long, totalCount As Long
        CountComponents swModel, visCount, totalCount
        If visCount < totalCount And visCount > 0 Then
            GenerateAssemblyReport swModel, True, visCount, totalCount
        Else
            GenerateAssemblyReport swModel, False, visCount, totalCount
        End If
        Exit Sub
    End If
    If swModel.GetType = swDocPART Then
        GeneratePartReport swModel
        Exit Sub
    End If
    MsgBox "Unknown document type!", vbExclamation
End Sub

' =====================================================
' COUNT HELPERS
' =====================================================
Function CountSelectedFaces(swSelMgr As SldWorks.SelectionMgr) As Long
    Dim c As Long, i As Long
    c = 0
    For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
        If swSelMgr.GetSelectedObjectType3(i, -1) = 2 Then c = c + 1
    Next i
    CountSelectedFaces = c
End Function

Function CountSelectedComponents(swSelMgr As SldWorks.SelectionMgr) As Long
    Dim c As Long, i As Long
    c = 0
    For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
        If swSelMgr.GetSelectedObjectType3(i, -1) = 20 Then c = c + 1
    Next i
    CountSelectedComponents = c
End Function

Sub CountComponents(swModel As SldWorks.ModelDoc2, ByRef vc As Long, ByRef tc As Long)
    Dim swAssy As SldWorks.AssemblyDoc, v As Variant, cmp As SldWorks.Component2, cm As SldWorks.ModelDoc2, i As Long
    Set swAssy = swModel: v = swAssy.GetComponents(False): vc = 0: tc = 0
    If IsEmpty(v) Then Exit Sub
    For i = 0 To UBound(v)
        Set cmp = v(i)
        If Not cmp.IsSuppressed And Not cmp.IsEnvelope Then
            Set cm = cmp.GetModelDoc2
            If Not cm Is Nothing Then
                If cm.GetType = swDocPART Then
                    tc = tc + 1
                    If Not cmp.IsHidden(True) Then vc = vc + 1
                End If
            End If
        End If
    Next i
End Sub

' =====================================================
' STRING HELPERS
' =====================================================
Function NStr(num As Double) As String
    NStr = Replace(Format(num, "0.0000"), ",", ".")
End Function

Function FmtF(num As Double) As String
    FmtF = Replace(Format(num, "0.00"), ",", ".")
End Function

Function FormatNum(num As Double) As String
    FormatNum = Format(num, "#,##0.00")
End Function

Function Enc(text As String) As String
    Dim r As String: r = text
    r = Replace(r, "&", "&amp;"): r = Replace(r, "<", "&lt;"): r = Replace(r, ">", "&gt;"): r = Replace(r, """", "&quot;")
    Enc = r
End Function

' =====================================================
' TAKE SCREENSHOT
' =====================================================
Function TakeScreenshot(swModel As SldWorks.ModelDoc2) As String
    On Error GoTo SSerr
    swModel.ShowNamedView2 "*Isometric", -1
    swModel.ViewZoomtofit2
    DoEvents
    Dim imgPath As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim swView As SldWorks.ModelView
    imgPath = Environ("TEMP") & "\sw_report_thumb.png"
    On Error Resume Next: Kill imgPath: On Error GoTo 0
    Set swView = swModel.ActiveView
    If Not swView Is Nothing Then
        On Error Resume Next: swView.SaveImage imgPath: On Error GoTo 0
    End If
    If fso.FileExists(imgPath) Then GoTo GotImage
    imgPath = Environ("TEMP") & "\sw_report_thumb.bmp"
    On Error Resume Next: Kill imgPath: On Error GoTo 0
    On Error Resume Next: swView.SaveImage imgPath: On Error GoTo 0
    If fso.FileExists(imgPath) Then GoTo GotImage
    imgPath = Environ("TEMP") & "\sw_report_thumb.bmp"
    On Error Resume Next: Kill imgPath: On Error GoTo 0
    On Error Resume Next
    Dim swExt As Object: Set swExt = swModel.Extension
    If Not swExt Is Nothing Then
        Dim errors As Long, warnings As Long
        swExt.SaveAs imgPath, 0, 0, Nothing, errors, warnings
    End If
    On Error GoTo 0
    If fso.FileExists(imgPath) Then GoTo GotImage
    TakeScreenshot = ""
    Exit Function
GotImage:
    TakeScreenshot = Base64FromFile(imgPath)
    On Error Resume Next: Kill imgPath: On Error GoTo 0
    Exit Function
SSerr:
    TakeScreenshot = ""
End Function

' =====================================================
' BASE64 FROM FILE
' =====================================================
Function Base64FromFile(filePath As String) As String
    On Error GoTo B64Err
    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1: stream.Open: stream.LoadFromFile filePath
    Dim binData As Variant: binData = stream.Read(-1): stream.Close
    Dim dom As Object: Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    Dim elem As Object: Set elem = dom.createElement("b64")
    elem.DataType = "bin.base64": elem.nodeTypedValue = binData
    Dim result As String
    result = Replace(elem.text, vbLf, ""): result = Replace(result, vbCr, "")
    Base64FromFile = result
    Exit Function
B64Err:
    Base64FromFile = ""
End Function

' =====================================================
' GET PREPARATION NOTES BASED ON MATERIAL
' =====================================================
Function GetPrepNotes(material As String) As String
    If Len(Trim(material)) = 0 Or material = "N/A" Then
        GetPrepNotes = "Clean surface, remove oil/grease before painting."
        Exit Function
    End If
    Dim m As String: m = LCase(material)
    If InStr(m, "stainless") > 0 Then
        GetPrepNotes = "Passivate surface, clean with solvent, apply etching primer if needed."
    ElseIf InStr(m, "alumin") > 0 Then
        GetPrepNotes = "Etch or abrade surface, clean with solvent, use self-etching primer."
    ElseIf InStr(m, "steel") > 0 Or InStr(m, "iron") > 0 Then
        GetPrepNotes = "Sandblast or grind to bare metal, clean, apply primer immediately."
    ElseIf InStr(m, "copper") > 0 Or InStr(m, "brass") > 0 Then
        GetPrepNotes = "Clean with degreaser, bright dip if needed, apply adhesion promoter."
    ElseIf InStr(m, "zinc") > 0 Or InStr(m, "galvan") > 0 Then
        GetPrepNotes = "Clean with ammonia-based solution, apply galvanized metal primer."
    ElseIf InStr(m, "plastic") > 0 Or InStr(m, "abs") > 0 Or InStr(m, "nylon") > 0 Then
        GetPrepNotes = "Clean with soap and water, scuff sand, apply plastic adhesion promoter."
    Else
        GetPrepNotes = "Clean surface, remove oil/grease before painting."
    End If
End Function

' =====================================================
' FACE TYPE NAME
' =====================================================
Function GetFaceTypeName(swFace As SldWorks.Face2) As String
    Dim swSurf As SldWorks.Surface
    Set swSurf = swFace.GetSurface
    If swSurf Is Nothing Then
        GetFaceTypeName = "Unknown"
        Exit Function
    End If
    If swSurf.IsPlane Then
        GetFaceTypeName = "Planar"
    ElseIf swSurf.IsCylinder Then
        GetFaceTypeName = "Cylindrical"
    ElseIf swSurf.IsCone Then
        GetFaceTypeName = "Conical"
    ElseIf swSurf.IsSphere Then
        GetFaceTypeName = "Spherical"
    ElseIf swSurf.IsTorus Then
        GetFaceTypeName = "Toroidal"
    Else
        GetFaceTypeName = "Complex"
    End If
End Function

' =====================================================
' GET FACE SIDE FROM NORMAL
' =====================================================
Function GetFaceSide(swFace As Face2) As String
    Dim nx As Double, ny As Double, nz As Double, gotNormal As Boolean
    Dim vNormal As Variant, swVec As Object
    gotNormal = False
    On Error Resume Next
    vNormal = swFace.Normal
    If IsArray(vNormal) Then
        If UBound(vNormal) >= 2 Then nx = CDbl(vNormal(0)): ny = CDbl(vNormal(1)): nz = CDbl(vNormal(2)): gotNormal = True
    End If
    On Error GoTo 0
    If Not gotNormal Then
        On Error Resume Next
        Set swVec = swFace.Normal
        If Not swVec Is Nothing Then
            vNormal = swVec.ArrayData
            If IsArray(vNormal) Then nx = CDbl(vNormal(0)): ny = CDbl(vNormal(1)): nz = CDbl(vNormal(2)): gotNormal = True
        End If
        On Error GoTo 0
    End If
    If Not gotNormal Then GetFaceSide = "Other": Exit Function
    Dim absX As Double, absY As Double, absZ As Double
    absX = Abs(nx): absY = Abs(ny): absZ = Abs(nz)
    Dim dominantAxis As Integer, isPositive As Boolean
    If absX >= absY And absX >= absZ Then
        dominantAxis = 1: isPositive = (nx > 0)
    ElseIf absY >= absX And absY >= absZ Then
        dominantAxis = 2: isPositive = (ny > 0)
    Else
        dominantAxis = 3: isPositive = (nz > 0)
    End If
    Dim frontAxis As Integer, topAxis As Integer, rightAxis As Integer
    If gFrontPlane = 1 Then
        frontAxis = 3: topAxis = 2: rightAxis = 1
    ElseIf gFrontPlane = 2 Then
        frontAxis = 2: topAxis = 3: rightAxis = 1
    Else
        frontAxis = 1: topAxis = 2: rightAxis = 3
    End If
    If dominantAxis = frontAxis Then
        GetFaceSide = IIf(isPositive, "Front", "Rear")
    ElseIf dominantAxis = topAxis Then
        GetFaceSide = IIf(isPositive, "Top", "Bottom")
    ElseIf dominantAxis = rightAxis Then
        GetFaceSide = IIf(isPositive, "Right", "Left")
    Else
        GetFaceSide = "Other"
    End If
End Function

' =====================================================
' FACE DETAIL STRING
' =====================================================
Function GetFaceDetailStr(swFace As SldWorks.Face2) As String
    Dim areaSqIn As Double, fType As String, side As String
    Dim vBox As Variant, d1 As Double, d2 As Double, d3 As Double, temp As Double
    areaSqIn = swFace.GetArea * 1550.0031
    fType = GetFaceTypeName(swFace): side = GetFaceSide(swFace)
    d1 = 0: d2 = 0: d3 = 0
    On Error Resume Next: vBox = swFace.GetBox: On Error GoTo 0
    If Not IsEmpty(vBox) Then
        If UBound(vBox) >= 5 Then
            d1 = (CDbl(vBox(3)) - CDbl(vBox(0))) * 39.3701
            d2 = (CDbl(vBox(4)) - CDbl(vBox(1))) * 39.3701
            d3 = (CDbl(vBox(5)) - CDbl(vBox(2))) * 39.3701
            If d1 < d2 Then temp = d1: d1 = d2: d2 = temp
            If d2 < d3 Then temp = d2: d2 = d3: d3 = temp
            If d1 < d2 Then temp = d1: d1 = d2: d2 = temp
        End If
    End If
    GetFaceDetailStr = NStr(areaSqIn) & "~" & NStr(d1) & "~" & NStr(d2) & "~" & NStr(d3) & "~" & fType & "~" & side
End Function

' =====================================================
' BUILD FORMULA PER SURFACE TYPE
' =====================================================
Sub BuildFaceFormula(d1 As Double, d2 As Double, d3 As Double, fType As String, areaSqIn As Double, ByRef formulaHtml As String, ByRef dimsHtml As String)
    Dim PS As String: PS = ChrW(960)
    Dim diameter As Double, height As Double, radius As Double, slantH As Double
    Dim majorR As Double, minorR As Double
    Select Case fType
    Case "Planar"
        dimsHtml = FmtF(d1) & " x " & FmtF(d2)
        formulaHtml = FmtF(d1) & " x " & FmtF(d2) & " = " & FormatNum(areaSqIn) & " sq in"
    Case "Cylindrical"
        If (d1 - d2) <= (d2 - d3) Then diameter = (d1 + d2) / 2: height = d3 Else diameter = (d2 + d3) / 2: height = d1
        dimsHtml = "Dia " & FmtF(diameter) & " x H " & FmtF(height)
        formulaHtml = PS & " x " & FmtF(diameter) & " x " & FmtF(height) & " = " & FormatNum(areaSqIn) & " sq in"
    Case "Conical"
        If (d1 - d2) <= (d2 - d3) Then diameter = (d1 + d2) / 2: height = d3 Else diameter = (d2 + d3) / 2: height = d1
        radius = diameter / 2
        If radius > 0 And height > 0 Then slantH = Sqr(radius * radius + height * height) Else slantH = 0
        dimsHtml = "Dia " & FmtF(diameter) & " x H " & FmtF(height)
        formulaHtml = PS & " x " & FmtF(radius) & " x " & FmtF(slantH) & " = " & FormatNum(areaSqIn) & " sq in"
    Case "Spherical"
        diameter = (d1 + d2 + d3) / 3: radius = diameter / 2
        dimsHtml = "Dia " & FmtF(diameter)
        formulaHtml = "4" & PS & " x " & FmtF(radius) & "<sup>2</sup> = " & FormatNum(areaSqIn) & " sq in"
    Case "Toroidal"
        majorR = (d1 - d3) / 2: minorR = d3 / 2
        If majorR < 0 Then majorR = 0
        dimsHtml = "R " & FmtF(majorR) & " x r " & FmtF(minorR)
        formulaHtml = "4" & PS & "<sup>2</sup> x " & FmtF(majorR) & " x " & FmtF(minorR) & " = " & FormatNum(areaSqIn) & " sq in"
    Case Else
        dimsHtml = FmtF(d1) & " x " & FmtF(d2) & " x " & FmtF(d3)
        formulaHtml = FormatNum(areaSqIn) & " sq in"
    End Select
End Sub

' =====================================================
' GAUGE LOOKUP
' =====================================================
Function GetGauge(thicknessIn As Double) As String
    Dim gData As Variant
    gData = Array(7, 0.1793, 8, 0.1644, 9, 0.1495, 10, 0.1345, 11, 0.1196, _
                  12, 0.1046, 13, 0.0897, 14, 0.0747, 15, 0.0673, 16, 0.0598, _
                  17, 0.0538, 18, 0.0478, 19, 0.0418, 20, 0.0359, 22, 0.0299, _
                  24, 0.0239, 26, 0.0179, 28, 0.0149, 30, 0.012)
    Dim bestG As Long, bestD As Double, i As Long
    bestG = 0: bestD = 999
    For i = 0 To UBound(gData) - 1 Step 2
        If Abs(thicknessIn - CDbl(gData(i + 1))) < bestD Then
            bestD = Abs(thicknessIn - CDbl(gData(i + 1))): bestG = CLng(gData(i))
        End If
    Next i
    If bestG > 0 Then GetGauge = CStr(bestG) & " ga" Else GetGauge = ""
End Function

' =====================================================
' GET PART MATERIAL INFO
' =====================================================
Function GetPartMaterialInfo(swCompModel As SldWorks.ModelDoc2) As String
    Dim matName As String, swFeat As SldWorks.Feature, featType As String
    Dim thickM As Double, thickIn As Double, gauge As String
    matName = ""
    On Error Resume Next: matName = swCompModel.GetMaterialPropertyName2(""): On Error GoTo 0
    If Len(Trim(matName)) = 0 Then
        On Error Resume Next: matName = swCompModel.GetCustomInfoValue("", "Material"): On Error GoTo 0
    End If
    If Len(Trim(matName)) = 0 Then matName = "N/A"
    On Error Resume Next
    Set swFeat = swCompModel.FirstFeature
    Do While Not swFeat Is Nothing
        featType = ""
        On Error Resume Next: featType = swFeat.GetTypeName2: On Error GoTo 0
        If StrComp(featType, "SheetMetal", vbTextCompare) = 0 Then
            thickM = 0
            On Error Resume Next
            Dim swDef As Object: Set swDef = swFeat.GetDefinition
            If Not swDef Is Nothing Then thickM = swDef.Thickness
            On Error GoTo 0
            If thickM > 0 Then
                thickIn = thickM * 39.3701
                gauge = GetGauge(thickIn)
                If Len(gauge) > 0 Then
                    GetPartMaterialInfo = matName & " - " & gauge & " (" & FmtF(thickIn) & " in)"
                Else
                    GetPartMaterialInfo = matName & " - " & FmtF(thickIn) & " in"
                End If
            Else
                GetPartMaterialInfo = matName & " (Sheet)"
            End If
            Exit Function
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
    On Error GoTo 0
    GetPartMaterialInfo = matName
End Function

' =====================================================
' GET COMPONENT DIMS
' =====================================================
Function GetComponentDims(swComp As SldWorks.Component2) As String
    Dim swCompModel As SldWorks.ModelDoc2, swPart As SldWorks.PartDoc
    Dim vBox As Variant, dx As Double, dy As Double, dz As Double, t As Double
    Set swCompModel = swComp.GetModelDoc2
    If swCompModel Is Nothing Then GetComponentDims = "": Exit Function
    If swCompModel.GetType <> swDocPART Then GetComponentDims = "": Exit Function
    Set swPart = swCompModel
    On Error Resume Next: vBox = swPart.GetPartBox(True): On Error GoTo 0
    If IsEmpty(vBox) Then GetComponentDims = "": Exit Function
    If UBound(vBox) < 5 Then GetComponentDims = "": Exit Function
    dx = Abs(CDbl(vBox(3)) - CDbl(vBox(0))) * 39.3701
    dy = Abs(CDbl(vBox(4)) - CDbl(vBox(1))) * 39.3701
    dz = Abs(CDbl(vBox(5)) - CDbl(vBox(2))) * 39.3701
    If dx < dy Then t = dx: dx = dy: dy = t
    If dy < dz Then t = dy: dy = dz: dz = t
    If dx < dy Then t = dx: dx = dy: dy = t
    GetComponentDims = FmtF(dx) & " x " & FmtF(dy) & " x " & FmtF(dz)
End Function

' =====================================================
' GET SINGLE PART DIMS
' =====================================================
Function GetSinglePartDims(swModel As SldWorks.ModelDoc2) As String
    On Error GoTo SPDErr
    Dim swPart As SldWorks.PartDoc, vBox As Variant
    Dim dx As Double, dy As Double, dz As Double, t As Double
    Set swPart = swModel
    vBox = swPart.GetPartBox(True)
    If IsEmpty(vBox) Then GoTo SPDErr
    If UBound(vBox) < 5 Then GoTo SPDErr
    dx = Abs(CDbl(vBox(3)) - CDbl(vBox(0))) * 39.3701
    dy = Abs(CDbl(vBox(4)) - CDbl(vBox(1))) * 39.3701
    dz = Abs(CDbl(vBox(5)) - CDbl(vBox(2))) * 39.3701
    If dx < dy Then t = dx: dx = dy: dy = t
    If dy < dz Then t = dy: dy = dz: dz = t
    If dx < dy Then t = dx: dx = dy: dy = t
    GetSinglePartDims = FmtF(dx) & " x " & FmtF(dy) & " x " & FmtF(dz)
    Exit Function
SPDErr:
    GetSinglePartDims = "N/A"
End Function

' =====================================================
' COMPONENT FACE DETAILS / AREA
' =====================================================
Function GetComponentFaceDetails(swComp As SldWorks.Component2) As String
    Dim cm As SldWorks.ModelDoc2, p As SldWorks.PartDoc, vb As Variant, vi As Variant, b As SldWorks.Body2, f As SldWorks.Face2, r As String, i As Long
    r = "": Set cm = swComp.GetModelDoc2
    If cm Is Nothing Then GetComponentFaceDetails = "": Exit Function
    If cm.GetType <> swDocPART Then GetComponentFaceDetails = "": Exit Function
    On Error Resume Next: vb = swComp.GetBodies3(swSolidBody, vi): On Error GoTo 0
    If IsEmpty(vb) Then Set p = cm: vb = p.GetBodies2(swSolidBody, True)
    If IsEmpty(vb) Then vb = p.GetBodies2(swSolidBody, False)
    If IsEmpty(vb) Then GetComponentFaceDetails = "": Exit Function
    For i = 0 To UBound(vb)
        Set b = vb(i)
        If Not b Is Nothing Then
            Set f = b.GetFirstFace
            Do While Not f Is Nothing
                If Len(r) > 0 Then r = r & "|"
                r = r & GetFaceDetailStr(f)
                Set f = f.GetNextFace
            Loop
        End If
    Next i
    GetComponentFaceDetails = r
End Function

Function GetComponentArea(swComp As SldWorks.Component2) As Double
    Dim cm As SldWorks.ModelDoc2, p As SldWorks.PartDoc, vb As Variant, vi As Variant, b As SldWorks.Body2, f As SldWorks.Face2, ta As Double, i As Long
    ta = 0: Set cm = swComp.GetModelDoc2
    If cm Is Nothing Then GetComponentArea = 0: Exit Function
    If cm.GetType <> swDocPART Then GetComponentArea = 0: Exit Function
    cm.EditRebuild3
    On Error Resume Next: vb = swComp.GetBodies3(swSolidBody, vi): On Error GoTo 0
    If IsEmpty(vb) Then Set p = cm: vb = p.GetBodies2(swSolidBody, True)
    If IsEmpty(vb) Then vb = p.GetBodies2(swSolidBody, False)
    If IsEmpty(vb) Then GetComponentArea = 0: Exit Function
    For i = 0 To UBound(vb)
        Set b = vb(i)
        If Not b Is Nothing Then
            Set f = b.GetFirstFace
            Do While Not f Is Nothing: ta = ta + f.GetArea: Set f = f.GetNextFace: Loop
        End If
    Next i
    GetComponentArea = ta
End Function

' =====================================================
' GENERATE REPORTS
' =====================================================
Sub GenerateFacesReport(swModel As SldWorks.ModelDoc2, swSelMgr As SldWorks.SelectionMgr, fc As Long)
    Dim f As SldWorks.Face2, i As Long, ta As Double, idx As Long
    Dim fd() As String: ReDim fd(fc - 1): ta = 0: idx = 0
    For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
        If swSelMgr.GetSelectedObjectType3(i, -1) = 2 Then
            Set f = swSelMgr.GetSelectedObject6(i, -1)
            If Not f Is Nothing Then fd(idx) = GetFaceDetailStr(f): ta = ta + f.GetArea: idx = idx + 1
        End If
    Next i
    Dim thumbB64 As String: thumbB64 = TakeScreenshot(swModel)
    SaveAndOpenHTML BuildFacesHTML(swModel.GetTitle, fc, fd, ta, thumbB64), "FacesReport.html"
End Sub

Sub GenerateSelectedComponentsReport(swModel As SldWorks.ModelDoc2, swSelMgr As SldWorks.SelectionMgr)
    Dim c As SldWorks.Component2, cm As SldWorks.ModelDoc2, i As Long, st As Long
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    For i = 1 To swSelMgr.GetSelectedObjectCount2(-1)
        st = swSelMgr.GetSelectedObjectType3(i, -1)
        If st = 20 Then
            Set c = swSelMgr.GetSelectedObject6(i, -1)
            If Not c Is Nothing And Not c.IsSuppressed Then
                Set cm = c.GetModelDoc2
                If Not cm Is Nothing Then
                    If cm.GetType = swDocPART Then AddPartToDict c, d
                End If
            End If
        End If
    Next i
    Dim thumbB64 As String: thumbB64 = TakeScreenshot(swModel)
    SaveAndOpenHTML BuildAssemblyHTML(swModel.GetTitle, d, "SELECTED COMPONENTS", thumbB64), "SelectedComponentsReport.html"
End Sub

Sub GenerateAssemblyReport(swModel As SldWorks.ModelDoc2, vo As Boolean, vc As Long, tc As Long)
    Dim a As SldWorks.AssemblyDoc, v As Variant, c As SldWorks.Component2, cm As SldWorks.ModelDoc2
    Dim i As Long, inc As Boolean, d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set a = swModel: v = a.GetComponents(False)
    If IsEmpty(v) Then MsgBox "No components found!", vbExclamation: Exit Sub
    For i = 0 To UBound(v)
        Set c = v(i)
        If Not c.IsSuppressed And Not c.IsEnvelope Then
            inc = IIf(vo, Not c.IsHidden(True), True)
            If inc Then
                Set cm = c.GetModelDoc2
                If Not cm Is Nothing Then
                    If cm.GetType = swDocPART Then AddPartToDict c, d
                End If
            End If
        End If
    Next i
    Dim ms As String
    If vo Then ms = "ISOLATED VIEW (" & vc & " of " & tc & " visible)" Else ms = "FULL ASSEMBLY (" & tc & " parts)"
    Dim thumbB64 As String: thumbB64 = TakeScreenshot(swModel)
    SaveAndOpenHTML BuildAssemblyHTML(swModel.GetTitle, d, ms, thumbB64), "AssemblyReport.html"
End Sub

Sub GeneratePartReport(swModel As SldWorks.ModelDoc2)
    Dim swPart As SldWorks.PartDoc, vBodies As Variant, swBody As SldWorks.Body2, swFace As SldWorks.Face2
    Dim totalArea As Double, faceCount As Long, faceDetailsStr As String, i As Long
    Dim partMaterial As String, partDims As String, thumbB64 As String
    swModel.EditRebuild3: Set swPart = swModel: vBodies = swPart.GetBodies2(swSolidBody, True)
    totalArea = 0: faceCount = 0: faceDetailsStr = ""
    If Not IsEmpty(vBodies) Then
        For i = 0 To UBound(vBodies)
            Set swBody = vBodies(i)
            If Not swBody Is Nothing Then
                Set swFace = swBody.GetFirstFace
                Do While Not swFace Is Nothing
                    totalArea = totalArea + swFace.GetArea: faceCount = faceCount + 1
                    If Len(faceDetailsStr) > 0 Then faceDetailsStr = faceDetailsStr & "|"
                    faceDetailsStr = faceDetailsStr & GetFaceDetailStr(swFace)
                    Set swFace = swFace.GetNextFace
                Loop
            End If
        Next i
    End If
    partMaterial = GetPartMaterialInfo(swModel)
    partDims = GetSinglePartDims(swModel)
    thumbB64 = TakeScreenshot(swModel)
    SaveAndOpenHTML BuildPartHTML(swModel.GetTitle, totalArea, faceCount, faceDetailsStr, partMaterial, partDims, thumbB64), "PartReport.html"
End Sub

' =====================================================
' ADD PART TO DICTIONARY
' =====================================================
Sub AddPartToDict(swComp As SldWorks.Component2, partDict As Object)
    Dim pk As String, pn As String, pa As Double, dp As Long, sp As Long
    pk = swComp.GetPathName: pn = swComp.Name2
    dp = InStrRev(pn, "-"): If dp > 0 Then If IsNumeric(Mid(pn, dp + 1)) Then pn = Left(pn, dp - 1)
    sp = InStrRev(pn, "/"): If sp > 0 Then pn = Mid(pn, sp + 1)
    If partDict.Exists(pk) Then
        Dim e As Variant: e = partDict(pk): e(1) = e(1) + 1: e(3) = e(2) * e(1): partDict(pk) = e
    Else
        pa = GetComponentArea(swComp)
        Dim n(6) As Variant
        n(0) = pn: n(1) = 1: n(2) = pa: n(3) = pa
        n(4) = GetComponentFaceDetails(swComp)
        n(5) = GetPartMaterialInfo(swComp.GetModelDoc2)
        n(6) = GetComponentDims(swComp)
        partDict.Add pk, n
    End If
End Sub

' =====================================================
' BUILD SIDE GROUP HTML
' =====================================================
Function BuildSideGroupHTML(sideName As String, sideFaces() As String, ByRef runningTotalSqIn As Double) As String
    Dim html As String, i As Long, parts() As String
    Dim areaSqIn As Double, sideTotalSqIn As Double, sideColor As String
    Select Case sideName
        Case "Front": sideColor = "#2196F3"
        Case "Rear": sideColor = "#1565C0"
        Case "Top": sideColor = "#4CAF50"
        Case "Bottom": sideColor = "#2E7D32"
        Case "Right": sideColor = "#FF9800"
        Case "Left": sideColor = "#E65100"
        Case Else: sideColor = "#9E9E9E"
    End Select
    sideTotalSqIn = 0
    html = "<div class='side-group' style='border-left:4px solid " & sideColor & ";margin-bottom:10px;border-radius:6px;overflow:hidden;border:1px solid #e0e0e0;'>"
    html = html & "<div class='side-header' style='background:" & sideColor & "20;padding:8px 15px;font-weight:700;font-size:13px;color:" & sideColor & ";'>" & sideName & " SIDE (" & (UBound(sideFaces) + 1) & " faces)</div>"
    html = html & "<table class='face-table'><thead><tr><th style='text-align:left'>Face</th><th>Type</th><th>Dimensions</th><th>Formula</th><th>Area (sq in)</th><th>Area (sq ft)</th></tr></thead><tbody>"
    For i = 0 To UBound(sideFaces)
        parts = Split(sideFaces(i), "~"): areaSqIn = Val(parts(0)): sideTotalSqIn = sideTotalSqIn + areaSqIn
        Dim fType As String, d1 As Double, d2 As Double, d3 As Double, fHtml As String, dHtml As String
        fType = parts(4): d1 = Val(parts(1)): d2 = Val(parts(2)): d3 = Val(parts(3))
        BuildFaceFormula d1, d2, d3, fType, areaSqIn, fHtml, dHtml
        html = html & "<tr><td>Face " & (i + 1) & "</td><td>" & fType & "</td><td>" & dHtml & "</td><td class='formula-cell'>" & fHtml & "</td><td>" & FormatNum(areaSqIn) & "</td><td>" & FormatNumber(areaSqIn / 144, 4) & "</td></tr>"
    Next i
    html = html & "</tbody></table>"
    html = html & "<div style='padding:8px 15px;background:" & sideColor & "10;font-weight:700;font-size:13px;border-top:1px solid " & sideColor & "30;'>" & sideName & " Total: " & FormatNum(sideTotalSqIn) & " sq in (" & FormatNumber(sideTotalSqIn / 144, 4) & " sq ft)</div></div>"
    runningTotalSqIn = runningTotalSqIn + sideTotalSqIn
    BuildSideGroupHTML = html
End Function

' =====================================================
' BUILD FACE BREAKDOWN HTML
' =====================================================
Function BuildFaceBreakdownHTML(faceDetailsStr As String, qty As Long, material As String) As String
    Dim html As String
    If Len(Trim(faceDetailsStr)) = 0 Then
        BuildFaceBreakdownHTML = "<div class='face-detail'><p class='no-data'>Face data not available.</p></div>": Exit Function
    End If
    html = "<div class='face-detail'>"
    Dim prepSuggestion As String: prepSuggestion = GetPrepNotes(material)
    html = html & "<div class='prep-notes'><div class='prep-label'>Surface Preparation:</div>"
    html = html & "<textarea class='prep-textarea' onclick='event.stopPropagation()' rows='2'>" & Enc(prepSuggestion) & "</textarea></div>"
    Dim allFaces() As String: allFaces = Split(faceDetailsStr, "|")
    Dim sideDict As Object: Set sideDict = CreateObject("Scripting.Dictionary")
    Dim i As Long, side As String
    For i = 0 To UBound(allFaces)
        side = Split(allFaces(i), "~")(5)
        If sideDict.Exists(side) Then sideDict(side) = sideDict(side) & "|" & allFaces(i) Else sideDict.Add side, allFaces(i)
    Next i
    Dim sideOrder As Variant: sideOrder = Array("Front", "Rear", "Top", "Bottom", "Right", "Left", "Other")
    Dim s As Variant, totalSqIn As Double: totalSqIn = 0
    For Each s In sideOrder
        If sideDict.Exists(CStr(s)) Then
            Dim sideFaces() As String: sideFaces = Split(sideDict(CStr(s)), "|")
            html = html & BuildSideGroupHTML(CStr(s), sideFaces, totalSqIn)
        End If
    Next s
    html = html & "<div class='formula'><div class='formula-label'>Calculation:</div>"
    For i = 0 To UBound(allFaces)
        Dim fp() As String: fp = Split(allFaces(i), "~")
        Dim fa As Double: fa = Val(fp(0)): Dim ft2 As String: ft2 = fp(4)
        Dim fd1 As Double: fd1 = Val(fp(1)): Dim fd2 As Double: fd2 = Val(fp(2)): Dim fd3 As Double: fd3 = Val(fp(3))
        Dim ffHtml As String, fdHtml As String
        BuildFaceFormula fd1, fd2, fd3, ft2, fa, ffHtml, fdHtml
        html = html & "<div class='formula-line'>Face " & (i + 1) & " (" & ft2 & "): " & ffHtml & "</div>"
    Next i
    html = html & "<div class='formula-sep'></div>"
    html = html & "<div class='formula-line'><strong>Part Total: " & FormatNum(totalSqIn) & " sq in (" & FormatNumber(totalSqIn / 144, 4) & " sq ft) each</strong></div>"
    If qty > 1 Then html = html & "<div class='formula-line'><strong>" & FormatNum(totalSqIn) & " x " & qty & " = " & FormatNum(totalSqIn * qty) & " sq in (" & FormatNumber(totalSqIn * qty / 144, 4) & " sq ft)</strong></div>"
    html = html & "</div></div>"
    BuildFaceBreakdownHTML = html
End Function

' =====================================================
' BUILD FACES HTML
' =====================================================
Function BuildFacesHTML(fileName As String, fc As Long, fd() As String, ta As Double, thumbB64 As String) As String
    Dim html As String, i As Long, parts() As String
    Dim areaSqIn As Double, fType As String, side As String, pct As Double
    html = GetHTMLStart("Selected Faces Report", fileName, "SELECTED FACES (" & fc & " faces)")
    If Len(thumbB64) > 0 Then html = html & "<div style='text-align:center;padding:10px 30px;background:#f0f5fa;'><img src='data:image/png;base64," & thumbB64 & "' style='max-width:300px;border-radius:8px;border:1px solid #ccc;'></div>"
    html = html & "<div class='summary'>"
    html = html & "<div class='card'><div class='label'>Faces</div><div class='value'>" & fc & "</div></div>"
    html = html & "<div class='card highlight'><div class='label'>Total Area</div><div class='value'>" & FormatNum(ta * 1550.0031) & "</div><div class='unit'>sq in</div></div>"
    html = html & "<div class='card highlight'><div class='label'>Total Area</div><div class='value'>" & FormatNumber(ta * 10.76391, 4) & "</div><div class='unit'>sq ft</div></div>"
    html = html & "</div>"
    html = html & "<table id='report-table'><thead><tr>"
    html = html & "<th onclick='sortTable(0)'>No</th><th onclick='sortTable(1)'>Face</th><th onclick='sortTable(2)'>Type</th><th onclick='sortTable(3)'>Side</th>"
    html = html & "<th onclick='sortTable(4)'>Area (sq in)</th><th onclick='sortTable(5)'>Area (sq ft)</th><th onclick='sortTable(6)'>%</th>"
    html = html & "</tr></thead><tbody>"
    Dim tSqIn As Double: tSqIn = ta * 1550.0031
    For i = 0 To fc - 1
        parts = Split(fd(i), "~"): areaSqIn = Val(parts(0)): fType = parts(4): side = parts(5)
        If tSqIn > 0 Then pct = (areaSqIn / tSqIn) * 100 Else pct = 0
        html = html & "<tr><td>" & (i + 1) & "</td><td>Face " & (i + 1) & "</td><td>" & fType & "</td>"
        html = html & "<td><span class='side-badge side-" & LCase(side) & "'>" & side & "</span></td>"
        html = html & "<td>" & FormatNum(areaSqIn) & "</td><td>" & FormatNumber(areaSqIn / 144, 4) & "</td><td>" & FormatNumber(pct, 1) & "%</td></tr>"
    Next i
    html = html & "<tr class='total-row'><td></td><td colspan='4'>TOTAL</td><td>" & FormatNum(tSqIn) & "</td><td>" & FormatNumber(ta * 10.76391, 4) & "</td><td>100%</td></tr>"
    html = html & "</tbody></table>"
    html = html & GetHTMLEnd()
    BuildFacesHTML = html
End Function

' =====================================================
' BUILD ASSEMBLY HTML
' =====================================================
Function BuildAssemblyHTML(fileName As String, partDict As Object, modeStr As String, thumbB64 As String) As String
    Dim html As String, keys As Variant, i As Long, partData As Variant
    Dim totalQty As Long, grandTotal As Double, pct As Double
    keys = partDict.keys: totalQty = 0: grandTotal = 0
    For i = 0 To UBound(keys): partData = partDict(keys(i)): totalQty = totalQty + CLng(partData(1)): grandTotal = grandTotal + CDbl(partData(3)): Next i
    html = GetHTMLStart("Assembly Report", fileName, modeStr)
    If Len(thumbB64) > 0 Then html = html & "<div style='text-align:center;padding:10px 30px;background:#f0f5fa;'><img src='data:image/png;base64," & thumbB64 & "' style='max-width:350px;border-radius:8px;border:1px solid #ccc;'></div>"
    html = html & "<div class='summary'>"
    html = html & "<div class='card'><div class='label'>Unique Parts</div><div class='value'>" & partDict.count & "</div></div>"
    html = html & "<div class='card'><div class='label'>Total Qty</div><div class='value'>" & totalQty & "</div></div>"
    html = html & "<div class='card highlight'><div class='label'>Grand Total</div><div class='value'>" & FormatNum(grandTotal * 1550.0031) & "</div><div class='unit'>sq in</div></div>"
    html = html & "<div class='card highlight'><div class='label'>Grand Total</div><div class='value'>" & FormatNumber(grandTotal * 10.76391, 4) & "</div><div class='unit'>sq ft</div></div>"
    html = html & "</div>"
    html = html & "<table id='report-table'><thead><tr>"
    html = html & "<th onclick='sortTable(0)'>No</th><th onclick='sortTable(1)'>Part Name</th><th onclick='sortTable(2)'>Material</th>"
    html = html & "<th onclick='sortTable(3)'>Size (in)</th><th onclick='sortTable(4)'>Qty</th><th onclick='sortTable(5)'>Each (sq in)</th>"
    html = html & "<th onclick='sortTable(6)'>Total (sq in)</th><th onclick='sortTable(7)'>Total (sq ft)</th><th onclick='sortTable(8)'>%</th><th>Paint Color</th>"
    html = html & "</tr></thead><tbody>"
    For i = 0 To UBound(keys)
        partData = partDict(keys(i))
        Dim did As String: did = "detail-" & (i + 1)
        If grandTotal > 0 Then pct = (CDbl(partData(3)) / grandTotal) * 100 Else pct = 0
        html = html & "<tr class='part-row' data-detail='" & did & "' onclick=""toggleDetail('" & did & "')"">"
        html = html & "<td>" & (i + 1) & "</td>"
        html = html & "<td>" & Enc(CStr(partData(0))) & " <span class='expand-btn' id='" & did & "-btn'>&#9654;</span></td>"
        html = html & "<td>" & Enc(CStr(partData(5))) & "</td><td>" & CStr(partData(6)) & "</td>"
        html = html & "<td>" & CStr(partData(1)) & "</td>"
        html = html & "<td>" & FormatNum(CDbl(partData(2)) * 1550.0031) & "</td>"
        html = html & "<td>" & FormatNum(CDbl(partData(3)) * 1550.0031) & "</td>"
        html = html & "<td>" & FormatNumber(CDbl(partData(3)) * 10.76391, 4) & "</td>"
        html = html & "<td>" & FormatNumber(pct, 1) & "%</td>"
        html = html & "<td><input type='text' class='paint-input' placeholder='e.g. RAL 7016' onclick='event.stopPropagation()'></td>"
        html = html & "</tr>"
        html = html & "<tr class='detail-row' id='" & did & "' style='display:none'>"
        html = html & "<td colspan='10'>" & BuildFaceBreakdownHTML(CStr(partData(4)), CLng(partData(1)), CStr(partData(5))) & "</td>"
        html = html & "</tr>"
    Next i
    html = html & "<tr class='total-row'><td></td><td>GRAND TOTAL</td><td></td><td></td><td>" & totalQty & "</td><td></td>"
    html = html & "<td>" & FormatNum(grandTotal * 1550.0031) & "</td>"
    html = html & "<td>" & FormatNumber(grandTotal * 10.76391, 4) & "</td>"
    html = html & "<td>100%</td><td></td></tr>"
    html = html & "</tbody></table>"
    html = html & "<div class='print-prep'><div class='print-prep-title'>Surface Preparation Notes</div>"
    html = html & "<table class='print-prep-table'><thead><tr><th>Part</th><th>Material</th><th>Preparation Instructions</th></tr></thead><tbody id='print-prep-body'>"
    For i = 0 To UBound(keys)
        partData = partDict(keys(i))
        Dim prepNote As String: prepNote = GetPrepNotes(CStr(partData(5)))
        html = html & "<tr><td>" & Enc(CStr(partData(0))) & "</td><td>" & Enc(CStr(partData(5))) & "</td><td class='print-prep-note'>" & Enc(prepNote) & "</td></tr>"
    Next i
    html = html & "</tbody></table></div>"
    html = html & GetHTMLEnd()
    BuildAssemblyHTML = html
End Function

' =====================================================
' BUILD PART HTML
' =====================================================
Function BuildPartHTML(fileName As String, totalArea As Double, faceCount As Long, faceDetailsStr As String, partMaterial As String, partDims As String, thumbB64 As String) As String
    Dim html As String
    html = GetHTMLStart("Part Report", fileName, "TOTAL SURFACE AREA")
    If Len(thumbB64) > 0 Then html = html & "<div style='text-align:center;padding:10px 30px;background:#f0f5fa;'><img src='data:image/png;base64," & thumbB64 & "' style='max-width:300px;border-radius:8px;border:1px solid #ccc;'></div>"
    html = html & "<div class='summary'>"
    html = html & "<div class='card'><div class='label'>Faces</div><div class='value'>" & faceCount & "</div></div>"
    html = html & "<div class='card'><div class='label'>Material</div><div class='value' style='font-size:14px'>" & Enc(partMaterial) & "</div></div>"
    html = html & "<div class='card'><div class='label'>Size (in)</div><div class='value' style='font-size:14px'>" & Enc(partDims) & "</div></div>"
    html = html & "<div class='card highlight'><div class='label'>Total Area</div><div class='value'>" & FormatNum(totalArea * 1550.0031) & "</div><div class='unit'>sq in</div></div>"
    html = html & "<div class='card highlight'><div class='label'>Total Area</div><div class='value'>" & FormatNumber(totalArea * 10.76391, 4) & "</div><div class='unit'>sq ft</div></div>"
    html = html & "</div>"
    html = html & "<table id='report-table'><thead><tr>"
    html = html & "<th>No</th><th>Part</th><th>Material</th><th>Size (in)</th><th>Faces</th><th>Area (sq in)</th><th>Area (sq ft)</th><th>Paint Color</th>"
    html = html & "</tr></thead><tbody>"
    Dim detailId As String: detailId = "detail-1"
    html = html & "<tr class='part-row' data-detail='" & detailId & "' onclick=""toggleDetail('" & detailId & "')"">"
    html = html & "<td>1</td><td>" & Enc(fileName) & " <span class='expand-btn' id='" & detailId & "-btn'>&#9654;</span></td>"
    html = html & "<td>" & Enc(partMaterial) & "</td><td>" & Enc(partDims) & "</td><td>" & faceCount & "</td>"
    html = html & "<td>" & FormatNum(totalArea * 1550.0031) & "</td><td>" & FormatNumber(totalArea * 10.76391, 4) & "</td>"
    html = html & "<td><input type='text' class='paint-input' placeholder='e.g. RAL 7016' onclick='event.stopPropagation()'></td></tr>"
    html = html & "<tr class='detail-row' id='" & detailId & "' style='display:none'>"
    html = html & "<td colspan='8'>" & BuildFaceBreakdownHTML(faceDetailsStr, 1, partMaterial) & "</td></tr>"
    html = html & "<tr class='total-row'><td></td><td>TOTAL</td><td></td><td></td><td>" & faceCount & "</td>"
    html = html & "<td>" & FormatNum(totalArea * 1550.0031) & "</td><td>" & FormatNumber(totalArea * 10.76391, 4) & "</td><td></td></tr>"
    html = html & "</tbody></table>"
    Dim prepNotePart As String: prepNotePart = GetPrepNotes(partMaterial)
    html = html & "<div class='print-prep'><div class='print-prep-title'>Surface Preparation Notes</div>"
    html = html & "<table class='print-prep-table'><thead><tr><th>Part</th><th>Material</th><th>Preparation Instructions</th></tr></thead><tbody id='print-prep-body'>"
    html = html & "<tr><td>" & Enc(fileName) & "</td><td>" & Enc(partMaterial) & "</td><td class='print-prep-note'>" & Enc(prepNotePart) & "</td></tr>"
    html = html & "</tbody></table></div>"
    html = html & GetHTMLEnd()
    BuildPartHTML = html
End Function

' =====================================================
' HTML START - EDITABLE COMPANY NAME + METADATA
' =====================================================
Function GetHTMLStart(title As String, fileName As String, modeStr As String) As String
    Dim h As String
    h = "<!DOCTYPE html><html><head><meta charset='UTF-8'><title>" & Enc(title) & "</title><style>"
    
    ' CSS Variables
    h = h & ":root{--primary:#0070C0;--primary-light:#00B4D8;--accent:#FF6B35;--bg-dark:#1a1a2e;--bg-card:#f5f5f5;--bg-highlight:#FFF5F0;--text-dark:#333;--text-muted:#666;--text-light:#888;--table-header:#2C3E50;--total-bg:#FFF3CD;--border:#eee;--hover:#e8f4fd;--detail-bg:#f8fbff;--radius:12px;--font-main:'Segoe UI',sans-serif;--font-mono:Consolas,monospace;--gold:#F59E0B;--gold-dark:#D97706;--gold-light:#FCD34D;--charcoal:#0F172A;--slate:#1E293B;}"
    
    ' Base
    h = h & "*{margin:0;padding:0;box-sizing:border-box;}"
    h = h & "body{font-family:var(--font-main);background:var(--bg-dark);padding:30px;min-height:100vh;}"
    h = h & ".container{background:#fff;border-radius:var(--radius);max-width:1050px;margin:0 auto;overflow:hidden;box-shadow:0 20px 60px rgba(0,0,0,0.3);}"
    
    ' *** NEW: COMPANY BRANDING STRIP - Gold on Dark Charcoal ***
    h = h & ".company-strip{background:linear-gradient(135deg,#0F172A 0%,#1a2744 100%);padding:22px 30px 18px;border-bottom:4px solid var(--gold);position:relative;}"
    h = h & ".company-strip::after{content:'';position:absolute;bottom:0;left:0;right:0;height:4px;background:linear-gradient(90deg,var(--gold),var(--gold-light),var(--gold));}"
    h = h & ".company-input-wrap{position:relative;}"
    h = h & ".company-input{width:100%;background:transparent;border:2px dashed transparent;border-bottom:2px dashed rgba(245,158,11,0.3);font-size:34px;font-weight:900;color:var(--gold);text-transform:uppercase;letter-spacing:4px;padding:4px 0;outline:none;transition:border-color 0.3s;}"
    h = h & ".company-input::placeholder{color:rgba(245,158,11,0.25);font-style:italic;font-weight:700;letter-spacing:3px;}"
    h = h & ".company-input:focus{border-bottom-color:var(--gold);border-bottom-style:solid;}"
    h = h & ".company-input:not(:placeholder-shown){border-bottom-color:var(--gold);border-bottom-style:solid;}"
    h = h & ".company-subtitle{color:#64748B;font-size:9px;letter-spacing:4px;text-transform:uppercase;margin-top:6px;}"
    

    ' *** NEW: METADATA EDIT BAR ***
    h = h & ".meta-bar{display:flex;gap:20px;padding:14px 30px;background:#F8FAFC;border-bottom:1px solid #E2E8F0;align-items:flex-end;flex-wrap:wrap;}"
    h = h & ".meta-field{display:flex;flex-direction:column;gap:3px;}"
    h = h & ".meta-field label{font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:1.5px;color:#94A3B8;}"
    h = h & ".meta-input{background:#fff;border:1px solid #E2E8F0;border-radius:6px;padding:6px 12px;font-size:13px;color:#1E293B;width:200px;outline:none;font-family:var(--font-main);transition:border-color 0.2s,box-shadow 0.2s;}"
    h = h & ".meta-input:focus{border-color:var(--gold);box-shadow:0 0 0 3px rgba(245,158,11,0.1);}"
    h = h & ".meta-input::placeholder{color:#CBD5E1;font-style:italic;}"
    h = h & ".meta-spacer{flex:1;}"
    h = h & ".meta-date{font-size:11px;color:#94A3B8;font-weight:500;line-height:28px;}"
    
    ' Report Header (blue gradient - below company strip)
    h = h & ".header{background:linear-gradient(135deg,var(--primary),var(--primary-light));color:#fff;padding:20px 30px;}"
    h = h & ".header h1{font-size:22px;margin-bottom:5px;}"
    h = h & ".header .file{font-size:13px;opacity:0.9;margin-bottom:8px;}"
    h = h & ".header .mode{display:inline-block;background:rgba(255,255,255,0.2);padding:5px 15px;border-radius:15px;font-size:11px;}"
    
    ' Toolbar
    h = h & ".toolbar{display:flex;gap:10px;padding:15px 30px;background:var(--bg-card);border-bottom:1px solid var(--border);align-items:center;flex-wrap:wrap;}"
    h = h & ".search-input{padding:8px 15px;border:1px solid var(--border);border-radius:8px;font-size:13px;width:220px;outline:none;}"
    h = h & ".search-input:focus{border-color:var(--primary);box-shadow:0 0 0 3px rgba(0,112,192,0.1);}"
    h = h & ".toolbar-spacer{flex:1;}"
    h = h & ".btn-export{padding:8px 20px;border:none;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;color:#fff;transition:opacity 0.2s;}"
    h = h & ".btn-export:hover{opacity:0.85;}"
    h = h & ".btn-csv{background:var(--primary);}"
    h = h & ".btn-pdf{background:var(--accent);}"
    
    ' Content
    h = h & ".content{padding:25px 30px;}"
    
    ' Summary
    h = h & ".summary{display:flex;gap:12px;margin-bottom:25px;flex-wrap:wrap;}"
    h = h & ".card{flex:1;min-width:100px;background:var(--bg-card);border-radius:10px;padding:15px;border-left:4px solid var(--primary);}"
    h = h & ".card.highlight{border-left-color:var(--accent);background:var(--bg-highlight);}"
    h = h & ".card .label{font-size:10px;color:var(--text-muted);text-transform:uppercase;letter-spacing:1px;}"
    h = h & ".card .value{font-size:20px;font-weight:700;color:var(--text-dark);margin-top:5px;}"
    h = h & ".card .unit{font-size:11px;color:var(--text-muted);}"
    
    ' Table
    h = h & "table{width:100%;border-collapse:collapse;font-size:13px;}"
    h = h & "thead th{background:var(--table-header);color:#fff;padding:10px 12px;text-align:right;font-size:11px;text-transform:uppercase;cursor:pointer;user-select:none;white-space:nowrap;}"
    h = h & "thead th:hover{background:#1a252f;}"
    h = h & "thead th:first-child{text-align:center;width:40px;}"
    h = h & "thead th:nth-child(2){text-align:left;}"
    h = h & "thead th:nth-child(3){text-align:left;}"
    h = h & "thead th:nth-child(4){text-align:left;}"
    h = h & "tbody td{padding:8px 12px;border-bottom:1px solid var(--border);text-align:right;font-family:var(--font-mono);font-size:12px;}"
    h = h & "tbody td:first-child{text-align:center;color:var(--text-light);font-family:var(--font-main);}"
    h = h & "tbody td:nth-child(2){text-align:left;font-family:var(--font-main);font-weight:500;}"
    h = h & "tbody td:nth-child(3){text-align:left;font-family:var(--font-main);font-size:11px;}"
    h = h & "tbody td:nth-child(4){text-align:left;font-family:var(--font-mono);font-size:11px;}"
    h = h & "tbody tr:nth-child(even){background:#f9f9f9;}"
    h = h & "tbody tr:hover{background:var(--hover);}"
    h = h & ".total-row td{background:var(--total-bg)!important;font-weight:700!important;border-top:2px solid #333;}"
    
    ' Side badges
    h = h & ".side-badge{display:inline-block;padding:2px 8px;border-radius:10px;font-size:10px;font-weight:700;color:#fff;text-transform:uppercase;}"
    h = h & ".side-front{background:#2196F3;}.side-rear{background:#1565C0;}.side-top{background:#4CAF50;}.side-bottom{background:#2E7D32;}.side-right{background:#FF9800;}.side-left{background:#E65100;}.side-other{background:#9E9E9E;}"
    
    ' Expandable
    h = h & ".part-row{cursor:pointer;}.part-row:hover{background:var(--hover)!important;}"
    h = h & ".expand-btn{font-size:10px;color:var(--primary);margin-left:6px;display:inline-block;}"
    h = h & ".detail-row td{padding:0!important;background:var(--detail-bg)!important;border:none!important;}"
    
    ' Prep notes
    h = h & ".prep-notes{margin:12px 15px;padding:10px 15px;background:#f0f7ff;border:1px solid #b8d4f0;border-radius:6px;}"
    h = h & ".prep-label{font-size:10px;text-transform:uppercase;color:#1a6bc4;font-weight:700;margin-bottom:4px;letter-spacing:0.5px;}"
    h = h & ".prep-textarea{width:100%;min-height:36px;padding:6px 10px;border:1px solid #ddd;border-radius:4px;font-family:var(--font-main);font-size:12px;color:#333;resize:vertical;outline:none;}"
    h = h & ".prep-textarea:focus{border-color:var(--primary);box-shadow:0 0 0 2px rgba(0,112,192,0.15);}"
    
    ' Face breakdown
    h = h & ".face-detail{padding:15px 20px;}"
    h = h & ".face-table{width:100%;font-size:12px;border-collapse:collapse;}"
    h = h & ".face-table th{background:#e8f0fe;padding:6px 12px;text-align:right;font-size:10px;text-transform:uppercase;color:#555;font-weight:600;}"
    h = h & ".face-table th:first-child{text-align:left;}"
    h = h & ".face-table td{padding:4px 12px;border-bottom:1px solid #e0e8f0;text-align:right;font-family:var(--font-mono);font-size:12px;}"
    h = h & ".face-table td:first-child{text-align:left;font-family:var(--font-main);color:var(--text-muted);}"
    h = h & ".no-data{color:var(--text-muted);font-style:italic;padding:10px;}"
    h = h & ".side-group{margin-bottom:10px;border-radius:6px;overflow:hidden;border:1px solid #e0e0e0;}"
    h = h & ".formula{margin-top:12px;padding:12px 15px;background:#fff9e6;border-left:3px solid #f0c040;border-radius:0 6px 6px 0;}"
    h = h & ".formula-label{font-size:10px;text-transform:uppercase;color:#b8960c;font-weight:700;margin-bottom:6px;letter-spacing:1px;}"
    h = h & ".formula-line{font-family:var(--font-mono);font-size:12px;color:#555;line-height:1.8;word-break:break-all;}"
    h = h & ".formula-line strong{color:var(--text-dark);}"
    h = h & ".formula-sep{border-top:1px dashed #ddd;margin:6px 0;}"
    
    ' Paint input
    h = h & ".paint-input{width:100%;min-width:90px;padding:4px 8px;border:1px solid #ddd;border-radius:4px;font-size:12px;font-family:var(--font-main);outline:none;color:var(--text-dark);}"
    h = h & ".paint-input:focus{border-color:var(--primary);box-shadow:0 0 0 2px rgba(0,112,192,0.15);}"
    h = h & ".paint-input::placeholder{color:#bbb;font-style:italic;}"
    
    ' Print-only
    h = h & ".print-prep{display:none;}"
    h = h & ".print-prep-title{font-size:14pt;font-weight:bold;color:#1a3a5c;margin-bottom:8px;border-bottom:2px solid #1a3a5c;padding-bottom:4px;}"
    h = h & ".print-prep-table{width:100%;border-collapse:collapse;font-size:9pt;margin-top:5px;}"
    h = h & ".print-prep-table th{background:#1a3a5c;color:#fff;padding:6px 10px;text-align:left;font-size:8pt;border:1px solid #0d2137;}"
    h = h & ".print-prep-table td{padding:5px 10px;border:1px solid #ccc;vertical-align:top;}"
    h = h & ".print-prep-table td:last-child{min-width:250px;}"
    
    ' Footer with company
    h = h & ".footer{padding:16px 30px;background:linear-gradient(135deg,#0F172A 0%,#1a2744 100%);display:flex;justify-content:space-between;align-items:center;border-top:3px solid var(--gold);}"
    h = h & ".footer-company{font-weight:900;color:var(--gold);font-size:15px;letter-spacing:2px;text-transform:uppercase;}"
    h = h & ".footer-meta{text-align:right;font-size:11px;color:#94A3B8;}"
    
    ' ======== PRINT STYLES ========
    h = h & "@media print{"
    h = h & "@page{margin:12mm 10mm;}"
    h = h & "body{background:#fff!important;padding:0!important;font-size:11pt;}"
    h = h & ".container{box-shadow:none!important;max-width:100%!important;border-radius:0!important;}"
    
    ' Company strip print - GOLD on DARK, very prominent
    h = h & ".company-strip{background:#0F172A!important;-webkit-print-color-adjust:exact;print-color-adjust:exact;padding:18px 25px!important;border-bottom:5px solid #F59E0B!important;}"
    h = h & ".company-input{border:none!important;font-size:28pt!important;font-weight:900!important;color:#F59E0B!important;letter-spacing:3px!important;background:transparent!important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}"

h = h & ".company-subtitle{color:#64748B!important;font-size:8pt!important;}"
    
    ' Meta bar print
    h = h & ".meta-bar{background:#fff!important;border-bottom:2px solid #F59E0B!important;padding:10px 25px!important;}"
    h = h & ".meta-input{border:1px solid #ccc!important;background:#fff!important;font-size:10pt!important;color:#333!important;box-shadow:none!important;}"
    h = h & ".meta-field label{color:#666!important;}"
    
    ' Header print
    h = h & ".header{background:#1a3a5c!important;color:#fff!important;padding:15px 25px!important;border-bottom:3px solid #0d2137!important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}"
    h = h & ".header h1{font-size:16pt!important;}"
    h = h & ".header .file{font-size:10pt!important;opacity:1!important;}"
    h = h & ".header .mode{font-size:8pt!important;background:rgba(255,255,255,0.25)!important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}"
    
    h = h & ".toolbar{display:none!important;}"
    h = h & ".expand-btn{display:none!important;}"
    h = h & ".detail-row{display:none!important;}"
    h = h & ".no-export{display:none!important;}"
    h = h & ".part-row{cursor:default!important;}"
    h = h & "thead th{cursor:default!important;}"
    h = h & ".content{padding:15px 25px!important;}"
    h = h & ".summary{display:flex!important;gap:0!important;margin-bottom:15px!important;border:2px solid #1a3a5c!important;border-radius:0!important;overflow:hidden!important;}"
    h = h & ".card{min-width:auto!important;padding:10px 15px!important;border-left:1px solid #ccc!important;border-radius:0!important;text-align:center!important;}"
    h = h & ".card:first-child{border-left:none!important;}"
    h = h & ".card.highlight{background:#fff3e0!important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}"
    h = h & ".card .label{font-size:7pt!important;}"
    h = h & ".card .value{font-size:14pt!important;margin-top:0!important;}"
    h = h & "table{font-size:9pt!important;border:2px solid #1a3a5c!important;}"
    h = h & "thead th{background:#1a3a5c!important;color:#fff!important;padding:6px 10px!important;font-size:8pt!important;border:1px solid #0d2137!important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}"
    h = h & "tbody td{padding:5px 10px!important;border:1px solid #ccc!important;font-size:9pt!important;}"
    h = h & "tbody tr:nth-child(even){background:#f5f8fb!important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}"
    h = h & ".total-row td{background:#FFF3CD!important;border:1px solid #c9b200!important;font-weight:700!important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}"
    h = h & ".side-badge{-webkit-print-color-adjust:exact;print-color-adjust:exact;}"
    h = h & ".paint-input{border:none!important;background:transparent!important;box-shadow:none!important;font-size:9pt!important;padding:0!important;width:100%!important;}"
    h = h & ".print-prep{display:block!important;margin-top:20px!important;padding:15px 25px!important;border-top:2px solid #1a3a5c!important;}"
    h = h & ".print-prep-table{-webkit-print-color-adjust:exact;print-color-adjust:exact;}"
    h = h & ".print-prep-table th{-webkit-print-color-adjust:exact;print-color-adjust:exact;}"
    
    ' Footer print
    h = h & ".footer{padding:14px 25px!important;border-top:4px solid #F59E0B!important;background:#0F172A!important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}"
    h = h & ".footer-company{font-size:12pt!important;font-weight:900!important;color:#F59E0B!important;-webkit-print-color-adjust:exact;print-color-adjust:exact;}"
    h = h & ".footer-meta{font-size:8pt!important;color:#94A3B8!important;}"
    h = h & "}"
    
    h = h & "</style></head><body>"
    h = h & "<div class='container'>"
    
    ' *** COMPANY BRANDING STRIP ***
    h = h & "<div class='company-strip'>"
    h = h & "<div class='company-input-wrap'>"
    h = h & "<input type='text' class='company-input' id='companyName' placeholder='CLICK HERE TO ENTER COMPANY NAME' oninput='updateFooter()'>"
    h = h & "</div>"
    h = h & "<div class='company-subtitle'>Surface Area Calculation Report</div>"
    h = h & "</div>"
    
    ' *** METADATA EDIT BAR ***
    h = h & "<div class='meta-bar'>"
    h = h & "<div class='meta-field'><label>Revision / Version</label><input type='text' class='meta-input' id='revisionInput' placeholder='e.g. Rev A, v2.1' oninput='updateFooter()'></div>"
    h = h & "<div class='meta-field'><label>Prepared By</label><input type='text' class='meta-input' id='authorInput' placeholder='e.g. John Smith' oninput='updateFooter()'></div>"
    h = h & "<div class='meta-field'><label>Project / PO Number</label><input type='text' class='meta-input' id='projectInput' placeholder='e.g. PO-2024-001'></div>"
    h = h & "<div class='meta-spacer'></div>"
    h = h & "<div class='meta-date'>" & Format(Now, "yyyy-mm-dd hh:mm") & "</div>"
    h = h & "</div>"
    
    ' Report Header (blue gradient)
    h = h & "<div class='header'><h1>" & Enc(title) & "</h1>"
    h = h & "<div class='file'>" & Enc(fileName) & "</div>"
    h = h & "<div class='mode'>" & Enc(modeStr) & "</div>"
    h = h & "</div>"
    
    ' Toolbar
    h = h & "<div class='toolbar'>"
    h = h & "<input type='text' class='search-input' placeholder='Search parts...' onkeyup='searchTable(this.value)'>"
    h = h & "<div class='toolbar-spacer'></div>"
    h = h & "<button class='btn-export btn-csv' onclick='exportExcel()'>Export Excel</button>"
    h = h & "<button class='btn-export btn-pdf' onclick='exportPDF()'>Export PDF</button>"
    h = h & "</div>"
    
    h = h & "<div class='content'>"
    GetHTMLStart = h
End Function

' =====================================================
' HTML END (JavaScript) - reads from editable inputs
' =====================================================
Function GetHTMLEnd() As String
    Dim h As String
    h = "</div>"
    
    ' Footer with dynamic company name
    h = h & "<div class='footer'>"
    h = h & "<div class='footer-company' id='footerCompany'></div>"
    h = h & "<div class='footer-meta'><span id='footerMeta'>"
    h = h & "Generated: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    h = h & "</span></div></div></div>"
    
    h = h & "<script>"
    
    ' --- Update footer dynamically ---
    h = h & "function updateFooter(){"
    h = h & "var c=document.getElementById('companyName').value;"
    h = h & "document.getElementById('footerCompany').textContent=c?c.toUpperCase():'';"
    h = h & "var r=document.getElementById('revisionInput').value;"
    h = h & "var a=document.getElementById('authorInput').value;"
    h = h & "var p=['Generated: " & Format(Now, "yyyy-mm-dd hh:mm:ss") & "'];"
    h = h & "if(r)p.push('Rev: '+r);"
    h = h & "if(a)p.push('By: '+a);"
    h = h & "document.getElementById('footerMeta').textContent=p.join(' | ');}"
    
    ' --- Sync prep notes ---
    h = h & "function syncPrepNotes(){"
    h = h & "var tas=document.querySelectorAll('.prep-textarea');"
    h = h & "var cells=document.querySelectorAll('.print-prep-note');"
    h = h & "for(var i=0;i<tas.length;i++){"
    h = h & "if(cells[i])cells[i].textContent=tas[i].value||tas[i].textContent;}"
    h = h & "updateFooter();}"
    h = h & "window.onbeforeprint=syncPrepNotes;"
    
    ' --- Toggle detail rows ---
    h = h & "function toggleDetail(id){"
    h = h & "var el=document.getElementById(id);"
    h = h & "if(!el)return;"
    h = h & "var open=el.style.display!=='none';"
    h = h & "el.style.display=open?'none':'table-row';"
    h = h & "var btn=document.getElementById(id+'-btn');"
    h = h & "if(btn)btn.innerHTML=open?'&#9654;':'&#9660;';}"
    
    ' --- Sort table (split into multiple lines) ---
    h = h & "function sortTable(col){"
    h = h & "var table=document.getElementById('report-table');"
    h = h & "if(!table)return;"
    h = h & "var tbody=table.querySelector('tbody');"
    h = h & "if(!tbody)return;"
    h = h & "var allRows=Array.from(tbody.children);"
    h = h & "var totalRow=null;var items=[];"
    h = h & "for(var i=0;i<allRows.length;i++){"
    h = h & "if(allRows[i].classList.contains('total-row')){totalRow=allRows[i];}"
    h = h & "else if(!allRows[i].classList.contains('detail-row')){"
    h = h & "var detId=allRows[i].getAttribute('data-detail');"
    h = h & "var detEl=detId?document.getElementById(detId):null;"
    h = h & "items.push({m:allRows[i],d:detEl});}}"
    h = h & "var key='sort-'+col;"
    h = h & "var asc=table.getAttribute(key)!=='a';"
    h = h & "table.setAttribute(key,asc?'a':'d');"
    h = h & "items.sort(function(a,b){"
    h = h & "var aT=a.m.cells[col]?a.m.cells[col].textContent.trim():'';"
    h = h & "var bT=b.m.cells[col]?b.m.cells[col].textContent.trim():'';"
    h = h & "var aN=parseFloat(aT.replace(/,/g,'').replace(/%/g,''));"
    h = h & "var bN=parseFloat(bT.replace(/,/g,'').replace(/%/g,''));"
    h = h & "if(!isNaN(aN)&&!isNaN(bN)){return asc?aN-bN:bN-aN;}"
    h = h & "return asc?aT.localeCompare(bT):bT.localeCompare(aT);});"
    h = h & "for(var i=0;i<items.length;i++){"
    h = h & "tbody.appendChild(items[i].m);"
    h = h & "if(items[i].d)tbody.appendChild(items[i].d);}"
    h = h & "if(totalRow)tbody.appendChild(totalRow);}"
    
    ' --- Search table (split into multiple lines) ---
    h = h & "function searchTable(q){"
    h = h & "var tbody=document.querySelector('#report-table tbody');"
    h = h & "if(!tbody)return;"
    h = h & "q=q.toLowerCase();"
    h = h & "for(var i=0;i<tbody.children.length;i++){"
    h = h & "var row=tbody.children[i];"
    h = h & "if(row.classList.contains('total-row'))continue;"
    h = h & "if(row.classList.contains('detail-row'))continue;"
    h = h & "var text=row.textContent.toLowerCase();"
    h = h & "var show=!q||text.indexOf(q)>=0;"
    h = h & "row.style.display=show?'':'none';"
    h = h & "var detId=row.getAttribute('data-detail');"
    h = h & "if(detId){var det=document.getElementById(detId);"
    h = h & "if(det&&!show)det.style.display='none';}}}"
    
    ' --- Excel export (split into many short lines) ---
    h = h & "function exportExcel(){"
    h = h & "syncPrepNotes();"
    h = h & "var table=document.getElementById('report-table');if(!table)return;"
    h = h & "var companyName=document.getElementById('companyName').value;"
    h = h & "var revision=document.getElementById('revisionInput').value;"
    h = h & "var author=document.getElementById('authorInput').value;"
    h = h & "var project=document.getElementById('projectInput').value;"
    h = h & "var hdr=document.querySelector('.header');"
    h = h & "var title=hdr?hdr.querySelector('h1').textContent:'';"
    h = h & "var fileEl=hdr?hdr.querySelector('.file'):null;"
    h = h & "var fileName=fileEl?fileEl.textContent:'';"
    h = h & "var modeEl=hdr?hdr.querySelector('.mode'):null;"
    h = h & "var modeStr=modeEl?modeEl.textContent:'';"
    h = h & "var dateStr='" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "';"
    h = h & "var q=String.fromCharCode(34);"
    h = h & "var NL=String.fromCharCode(10);var xh=[];"
    h = h & "xh.push('<html xmlns:o='+q+'urn:schemas-microsoft-com:office:office'+q"
    h = h & "+' xmlns:x='+q+'urn:schemas-microsoft-com:office:excel'+q+'>');"
    h = h & "xh.push('<head><meta charset='+q+'UTF-8'+q+'>');"
    h = h & "xh.push('<!--[if gte mso 9]><xml><x:ExcelWorkbook>'"
    h = h & "+'<x:ExcelWorksheets><x:ExcelWorksheet>'"
    h = h & "+'<x:Name>Report</x:Name>'"
    h = h & "+'<x:WorksheetOptions><x:DisplayGridlines/>'"
    h = h & "+'</x:WorksheetOptions></x:ExcelWorksheet>'"
    h = h & "+'</x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->');"
    h = h & "xh.push('<style>');"
    h = h & "xh.push('td,th{padding:6px 12px;font-family:Calibri,sans-serif;font-size:11pt;}');"
    h = h & "xh.push('table{border-collapse:collapse;width:100%;}');"
    h = h & "xh.push('</style></head><body>');"
    
    ' Company name in Excel
    h = h & "if(companyName){"
    h = h & "xh.push('<table style='+q+'width:100%;border:none;'+q+'>');"
    h = h & "xh.push('<tr><td style='+q"
    h = h & "+'font-size:26pt;font-weight:900;color:#D97706;padding:14px 0 6px 0;'"
    h = h & "+'letter-spacing:3px;text-transform:uppercase;border-bottom:4px solid #D97706;'"
    h = h & "+q+'>'+companyName.toUpperCase()+'</td></tr>');"
    h = h & "xh.push('<tr><td style='+q"
    h = h & "+'font-size:8pt;color:#999;letter-spacing:2px;padding:2px 0 10px 0;'"
    h = h & "+q+'>SURFACE AREA REPORT</td></tr>');"
    h = h & "xh.push('</table>');}"
    
    ' Meta info in Excel
    h = h & "xh.push('<table style='+q+'width:100%;border:none;margin-bottom:15px;'+q+'>');"
    h = h & "xh.push('<tr><td style='+q"
    h = h & "+'font-size:18pt;font-weight:bold;color:#1a3a5c;padding:8px 0;'"
    h = h & "+'border-bottom:2px solid #1a3a5c;'+q+'>'+title+'</td></tr>');"
    h = h & "xh.push('<tr><td style='+q+'font-size:11pt;color:#555;padding:5px 0;'+q+'>'+fileName+'</td></tr>');"
    h = h & "xh.push('<tr><td style='+q+'font-size:9pt;color:#888;padding:3px 0;'+q+'>'+modeStr+'</td></tr>');"
    h = h & "var mP=[];"
    h = h & "if(revision)mP.push('Revision: '+revision);"
    h = h & "if(author)mP.push('Prepared by: '+author);"
    h = h & "if(project)mP.push('Project: '+project);"
    h = h & "if(mP.length>0)xh.push('<tr><td style='+q"
    h = h & "+'font-size:9pt;color:#666;padding:3px 0;border-bottom:1px solid #ccc;'"
    h = h & "+q+'>'+mP.join(' | ')+'</td></tr>');"
    h = h & "xh.push('</table>');"
    
    ' Data table in Excel
    h = h & "xh.push('<table border='+q+'1'+q+' cellpadding='+q+'6'+q"
    h = h & "+' cellspacing='+q+'0'+q+' style='+q"
    h = h & "+'border-collapse:collapse;border:2px solid #1a3a5c;'+q+'>');"
    h = h & "var prepSection=document.querySelector('.print-prep');"
    h = h & "var hasPrep=prepSection?true:false;"
    h = h & "var mainTbody=table.querySelector('tbody');"
    h = h & "var mainThead=table.querySelector('thead');"
    h = h & "var rows=table.querySelectorAll('tr');"
    h = h & "for(var i=0;i<rows.length;i++){"
    h = h & "if(rows[i].classList.contains('detail-row'))continue;"
    h = h & "var pr=rows[i].parentElement;"
    h = h & "if(pr!==mainTbody&&pr!==mainThead)continue;"
    h = h & "var cells=rows[i].querySelectorAll('th,td');"
    h = h & "var isHeader=rows[i].parentElement.tagName==='THEAD';"
    h = h & "var isTotal=rows[i].classList.contains('total-row');"
    h = h & "xh.push('<tr>');"
    h = h & "for(var j=0;j<cells.length;j++){"
    h = h & "if(cells[j].classList.contains('no-export'))continue;"
    h = h & "var tag=isHeader?'th':'td';"
    h = h & "var txt=cells[j].textContent.trim();"
    h = h & "var pIn=cells[j].querySelector('.paint-input');"
    h = h & "if(pIn)txt=pIn.value||'';"
    h = h & "var stl='';"
    h = h & "if(isHeader){stl='background-color:#1a3a5c;color:#fff;'"
    h = h & "+'font-weight:bold;font-size:9pt;text-align:center;border:1px solid #0d2137;';}"
    h = h & "else if(isTotal){stl='background-color:#FFF3CD;'"
    h = h & "+'font-weight:bold;font-size:10pt;border:1px solid #c9b200;';}"
    h = h & "else{stl='border:1px solid #ccc;font-size:10pt;';"
    h = h & "if(j===0)stl+='text-align:center;color:#888;';"
    h = h & "else if(j===1)stl+='text-align:left;font-weight:600;';"
    h = h & "else if(j===2||j===3)stl+='text-align:left;';"
    h = h & "else stl+='text-align:right;font-family:Consolas,monospace;';}"
    h = h & "if(i%2===1&&!isHeader&&!isTotal)stl+='background-color:#f5f8fb;';"
    h = h & "xh.push('<'+tag+' style='+q+stl+q+'>'+txt+'</'+tag+'>');}"
    
    ' Prep notes column
    h = h & "if(!isTotal&&hasPrep){"
    h = h & "if(isHeader){xh.push('<th style='+q"
    h = h & "+'background-color:#1a3a5c;color:#fff;font-weight:bold;'"
    h = h & "+'font-size:9pt;text-align:center;border:1px solid #0d2137;'"
    h = h & "+q+'>Prep Notes</th>');}"
    h = h & "else{var detId=rows[i].getAttribute('data-detail');"
    h = h & "var prepText='';"
    h = h & "if(detId){var det=document.getElementById(detId);"
    h = h & "if(det){var ta=det.querySelector('.prep-textarea');"
    h = h & "if(ta)prepText=ta.value;}}"
    h = h & "xh.push('<td style='+q"
    h = h & "+'border:1px solid #ccc;font-size:9pt;text-align:left;max-width:300px;'"
    h = h & "+q+'>'+prepText+'</td>');}}"
    h = h & "xh.push('</tr>');}"
    h = h & "xh.push('</table>');"
    
    ' Footer in Excel
    h = h & "xh.push('<table style='+q+'width:100%;border:none;margin-top:10px;'+q+'>');"
    h = h & "if(companyName){xh.push('<tr><td colspan=2 style='+q"
    h = h & "+'font-size:12pt;font-weight:900;color:#D97706;padding:8px 0;'"
    h = h & "+'border-top:4px solid #D97706;letter-spacing:2px;text-transform:uppercase;'"
    h = h & "+q+'>'+companyName.toUpperCase()+'</td></tr>');}"
    h = h & "var fP=[dateStr];"
    h = h & "if(revision)fP.push('Rev: '+revision);"
    h = h & "if(author)fP.push('By: '+author);"
    h = h & "xh.push('<tr><td style='+q+'font-size:8pt;color:#888;'+q"
    h = h & "+'>'+fP.join(' | ')+'</td>'"
    h = h & "+'<td style='+q+'font-size:8pt;color:#888;text-align:right;'+q"
    h = h & "+'>Surface Area Report</td></tr>');"
    h = h & "xh.push('</table></body></html>');"
    h = h & "var content=xh.join(NL);"
    h = h & "var blob=new Blob([content],{type:'application/vnd.ms-excel;charset=utf-8'});"
    h = h & "var a=document.createElement('a');"
    h = h & "a.href=URL.createObjectURL(blob);"
    h = h & "a.download='SurfaceAreaReport.xls';"
    h = h & "document.body.appendChild(a);"
    h = h & "a.click();document.body.removeChild(a);}"
    
    ' PDF
    h = h & "function exportPDF(){window.print();}"
    
    h = h & "</script></body></html>"
    GetHTMLEnd = h
End Function

' =====================================================
' SAVE AND OPEN
' =====================================================
Sub SaveAndOpenHTML(html As String, fileName As String)
    Dim fso As Object, ts As Object, filePath As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    filePath = Environ("TEMP") & "\" & fileName
    Set ts = fso.CreateTextFile(filePath, True, True)
    ts.Write html: ts.Close
    Shell "explorer.exe """ & filePath & """", vbNormalFocus
End Sub
