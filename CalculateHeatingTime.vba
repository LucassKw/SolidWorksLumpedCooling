Option Explicit

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swPart As SldWorks.PartDoc
Dim swSelMgr As SldWorks.SelectionMgr
Dim swFace As SldWorks.Face2
Dim swMassProp As SldWorks.MassProperty
Dim vMatDBs As Variant
Dim sMatName As String, sMatDB As String, sMatDBPath As String
Dim dSpecificHeat As Double
Dim dDensity As Double, dVolume As Double, dArea As Double, dMass As Double
Dim dTi As Double, dTf As Double, dTinf As Double, h As Double
Dim iCoolantType As Integer
Dim dTime As Double
Dim i As Integer, j As Integer
Dim bRet As Boolean

' XML Objects
Dim xmlDoc As Object
Dim xmlMatNode As Object
Dim xmlPropNode As Object
Dim sXPath As String

Sub main()
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    ' Open and check part
    If swModel Is Nothing Then
        MsgBox "No active document."
        Exit Sub
    End If
    
    If swModel.GetType <> swDocPART Then
        MsgBox "Active document is not a part."
        Exit Sub
    End If
    
    Set swPart = swModel
    Set swSelMgr = swModel.SelectionManager
    
    ' 1. Calculate Selected Surface Area
    Dim nSelCount As Integer
    nSelCount = swSelMgr.GetSelectedObjectCount2(-1)
    dArea = 0
    
    If nSelCount = 0 Then
        MsgBox "Please select at least one face representing the cooling surface."
        Exit Sub
    End If
    
    For i = 1 To nSelCount
        If swSelMgr.GetSelectedObjectType3(i, -1) = swSelFACES Then
            Set swFace = swSelMgr.GetSelectedObject6(i, -1)
            dArea = dArea + swFace.GetArea
        End If
    Next i
    
    ' 2. Get Mass Properties
    Set swMassProp = swModel.Extension.CreateMassProperty
    dVolume = swMassProp.Volume
    
    dDensity = swModel.GetUserPreferenceDoubleValue(swUserPreferenceDoubleValue_e.swMaterialPropertyDensity)
    
    dMass = dVolume * dDensity
    
    If dMass = 0 Then
        MsgBox "Mass is zero. Ensure material is applied and geometry is valid."
        Exit Sub
    End If
    
    ' 3. Get Specific Heat (Cp) from XML
    If Not GetSpecificHeat(dSpecificHeat) Then
        MsgBox "Could not retrieve Specific Heat for current material."
        Exit Sub
    End If
    
    ' 4. User Inputs
    Dim sInput As String
    
    ' Initial Temperature
    sInput = InputBox("Enter Initial Temperature (Ti) in °C:", "Input Ti", "200")
    If sInput = "" Then Exit Sub
    dTi = CDbl(sInput)
    
    ' Final Temperature
    sInput = InputBox("Enter Target Final Temperature (Tf) in °C:", "Input Tf", "50")
    If sInput = "" Then Exit Sub
    dTf = CDbl(sInput)
    
    ' Coolant Temperature
    sInput = InputBox("Enter Coolant/Environment Temperature (Tinf) in °C:", "Input Tinf", "20")
    If sInput = "" Then Exit Sub
    dTinf = CDbl(sInput)

    ' --- Validation Checks ---
    If dTf >= dTi Then
        MsgBox "Error: Final temperature must be lower than Initial temperature. This macro is for Cooling scenarios only."
        Exit Sub
    End If
    
    If dTinf >= dTf Then
        MsgBox "Error: Coolant temperature (" & dTinf & ") is hotter than or equal to the desired final temperature (" & dTf & ")." & vbCrLf & _
               "Cooling to this target is impossible."
        Exit Sub
    End If
    
    ' Coolant Type Selection
    Dim sPrompt As String
    sPrompt = "Select Coolant Type (Enter Number):" & vbCrLf & _
              "1. Air (natural) [h=10]" & vbCrLf & _
              "2. Air (forced) [h=75]" & vbCrLf & _
              "3. Water [h=2000]" & vbCrLf & _
              "4. Water-glycol [h=1000]" & vbCrLf & _
              "5. Oil [h=200]" & vbCrLf & _
              "6. Boiling coolant [h=10000]"
              
    sInput = InputBox(sPrompt, "Select Coolant", "3")
    If sInput = "" Then Exit Sub
    iCoolantType = CInt(sInput)
    
    Select Case iCoolantType
        Case 1: h = 10
        Case 2: h = 75
        Case 3: h = 2000
        Case 4: h = 1000
        Case 5: h = 200
        Case 6: h = 10000
        Case Else
            MsgBox "Invalid selection. Defaulting to Air (natural) h=10."
            h = 10
    End Select
    
    ' 5. Calculate Time
    ' Formula: t = - (rho * V * Cp) / (h * A) * ln( (Tf - Tinf) / (Ti - Tinf) )
    
    Dim dTimeConst As Double
    dTimeConst = (dMass * dSpecificHeat) / (h * dArea)
    
    Dim dTempRatio As Double

    If Abs(dTi - dTinf) < 0.0001 Then
        MsgBox "Initial temperature is equal to coolant temperature. No heat transfer."
        Exit Sub
    End If
    
    dTempRatio = (dTf - dTinf) / (dTi - dTinf)
    
    If dTempRatio <= 0 Then
        MsgBox "Target temperature is physically impossible to reach with this coolant temperature (requires crossing the asymptote)."
        Exit Sub
    End If
    
    dTime = -1 * dTimeConst * Log(dTempRatio)
    
    ' 6. Output Results
    MsgBox "Calculation Results:" & vbCrLf & _
           "-------------------" & vbCrLf & _
           "Material: " & sMatName & vbCrLf & _
           "Specific Heat: " & dSpecificHeat & " J/kg.K" & vbCrLf & _
           "Total Mass: " & Format(dMass, "0.000") & " kg" & vbCrLf & _
           "Surface Area: " & Format(dArea, "0.0000") & " m^2" & vbCrLf & _
           "h coefficient: " & h & " W/m^2.K" & vbCrLf & _
           "-------------------" & vbCrLf & _
           "Time Required: " & Format(dTime, "0.00") & " seconds (" & Format(dTime / 60, "0.00") & " min)"
    
End Sub

Function GetSpecificHeat(ByRef dVal As Double) As Boolean
    sMatName = swPart.GetMaterialPropertyName2("", sMatDB)
    
    If sMatName = "" Then
        MsgBox "No material applied."
        GetSpecificHeat = False
        Exit Function
    End If
    
    vMatDBs = swApp.GetMaterialDatabases
    sMatDBPath = ""
    
    Dim k As Integer
    If Not IsEmpty(vMatDBs) Then
        For k = 0 To UBound(vMatDBs)
            If InStr(1, vMatDBs(k), sMatDB, vbTextCompare) > 0 Then
                sMatDBPath = vMatDBs(k)
                Exit For
            End If
        Next k
    End If
    
    If sMatDBPath = "" Then
        MsgBox "Database path not found."
        GetSpecificHeat = False
        Exit Function
    End If
    
    On Error Resume Next
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    If Err.Number <> 0 Then
        MsgBox "MSXML Error"
        GetSpecificHeat = False
        Exit Function
    End If
    On Error GoTo 0
    
    xmlDoc.async = False
    xmlDoc.Load sMatDBPath
    
    sXPath = "//material[@name='" & sMatName & "']"
    Set xmlMatNode = xmlDoc.SelectSingleNode(sXPath)
    
    If xmlMatNode Is Nothing Then
        MsgBox "Material node not found in XML."
        GetSpecificHeat = False
        Exit Function
    End If
    
    Dim found As Boolean: found = False
    Set xmlPropNode = xmlMatNode.SelectSingleNode(".//specific_heat")
    
    If Not xmlPropNode Is Nothing Then
        If Not xmlPropNode.Attributes.getNamedItem("value") Is Nothing Then
            dVal = CDbl(xmlPropNode.Attributes.getNamedItem("value").Text)
            found = True
        End If
    End If
    
    If Not found Then
        Set xmlPropNode = xmlMatNode.SelectSingleNode(".//C")
        If Not xmlPropNode Is Nothing Then
             If Not xmlPropNode.Attributes.getNamedItem("value") Is Nothing Then
                dVal = CDbl(xmlPropNode.Attributes.getNamedItem("value").Text)
                found = True
            End If
        End If
    End If
    
    GetSpecificHeat = found
End Function
