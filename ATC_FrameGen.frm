VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ATC_FrameGen 
   Caption         =   "ATC Steel Frame Generator"
   ClientHeight    =   10500
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8805
   OleObjectBlob   =   "ATC_FrameGen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ATC_FrameGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

 Public PartDirectoryPath As String
 Public WeldmentPath  As String               ' Directory Location of the Weldment Library
 Public ReturnVal As String                   ' Generic Return Value for Function Call, Not Currently Used for Evaluation
 Public MainTubeHeight As Double              ' Main Rail Tube Height
 Public OverChannel As String                 ' C-Channel Used in Steel Trailers where Main Rail / Tongue Passes Thru Channel
 Public InsideChannel As String               ' C-Channel Used as Cross Members
 Public SWAxleClearance As Double             '
 Public SWRailOverlap As Double               '
 Public SWHeaderTubeHeight As Double          '
 Public SWHeaderTubeWidth As Double           '
 Public SWFooterTubeHeight As Double          '
 Public SWFrontVerticalTubeHeight As Double   ' Side Wall Vertical Tube Height (1st Member of Wall)
 Public RoofTubeWidth As Double               '
 Public RoofTubeHeight As Double              '
 Public RampTubeHeight As Double              '
 Public RampTubeWidth As Double               '
 Public TrailerTotalHeight As Double               ' Trailer Height for Door Size, etc.
 Public AxleCL As Double                      ' Axls Center Line Location From Front of Frame
 Public FW_Offset As Double                   ' Front Wall Offset (along the length) *** These may be reversed ****
 Public SW_Offset As Double                   ' Side Wall Offset (laterial length)   *** These may be reversed ****
 Public Theta As Double                       ' Angle of the Tongue
 Public TGL As Double                         ' Tongue Gusset Length
 Public P1x As Double
 Public P1y As Double
 Public P2x As Double
 Public P2y As Double
 Public Dist As Double

Function WheelWellHeight(WWH As Double)
 
      If AxleRating.Value = "3500 lb 0 Deg" Then
         If FrameMember.Value = "Tube 2.0 x 5.0 x 11ga" Then
            WWH = 11
         ElseIf FrameMember.Value = "Tube 2.0 x 8.0 x 11ga" Then
            WWH = 8
         End If
      ElseIf AxleRating.Value = "5200 lb 10 Deg up" Then
         If FrameMember.Value = "Tube 2.0 x 5.0 x 11ga" Then
            WWH = 13
         ElseIf FrameMember.Value = "Tube 2.0 x 8.0 x 11ga" Then
            WWH = 10
         End If
      ElseIf AxleRating.Value = "6000 lb 10 Deg up" Then
         WWH = 14
      ElseIf AxleRating.Value = "6000 lb 22.5 Deg up" Then
         WWH = 12
      ElseIf AxleRating.Value = "7000 lb 10 Deg up" Then
         WWH = 14
      ElseIf AxleRating.Value = "7000 lb 22.5 Deg up" Then
         WWH = 12
      ElseIf AxleRating.Value = "8000 lb 22.5 Deg up" Then
         WWH = 12
      ElseIf AxleRating.Value = "10000 lb 22.5 Deg up" Then
         WWH = 12
      End If
      
End Function

Function SideWallTubeSize()
 
 If NoseStructure.Value = "2ft Wedge" Then     ' **************************** 2 ft Wedge ************************************
    FW_Offset = 0
    SW_Offset = 0
 ElseIf NoseStructure.Value = "4ft Wedge" Then ' **************************** 4 ft Wedge ************************************
    FW_Offset = 0
    SW_Offset = 0
 ElseIf NoseStructure.Value = "6ft Wedge" Then ' **************************** 6 ft Wedge ************************************
    FW_Offset = 0
    SW_Offset = 0
 Else
    FW_Offset = 5
    SW_Offset = 4
 End If

      If TrailerType.Value = "Motiv Raven Lite Cargo (Steel)" Then ' ********************* Wheel Over Side Wall Clearance **********************************
         SWAxleClearance = 57.5                          ' Raven: Dual Axle Wheel Clearance
         SWRailOverlap = 18                              ' Raven: Main Rail Overlap
         SWHeaderTubeHeight = 1.5
         SWHeaderTubeWidth = 1#
         SWFooterTubeHeight = 1.5
         SWFrontVerticalTubeHeight = 1.5
         'TrailerHeight = TrailerHeight.Value + 2.5      ' Add 2 Inches to Wall Height, Interior Wall Height Specified in Pull Down
         TrailerTotalHeight = TrailerHeight.Value + 2.5      ' Add 2 Inches to Wall Height, Interior Wall Height Specified in Pull Down
         
      ElseIf TrailerType.Value = "Motiv MSX Series" Then
         SWHeaderTubeHeight = 3#
         SWHeaderTubeWidth = 1#
         SWFooterTubeHeight = 1.5
         SWFrontVerticalTubeHeight = 3#
         'TrailerHeight.Value = TrailerHeight.Value + 2.5    ' Add 2.5 Inches to Wall Height, Interior Wall Height Specified in Pull Down
         TrailerTotalHeight = TrailerHeight.Value + 2.5      ' Add 2 Inches to Wall Height, Interior Wall Height Specified in Pull Down
         If AxleSpacing.Value = "Standard Axle Spacing" Then       ' ********************* Wheel Over Side Wall Clearance **********************************
            If NoOfAxles.Value = "Dual Axle" Then
               SWAxleClearance = 72                       ' Axle Wheel Clearance
               SWRailOverlap = 18                         ' Wall Rail Overlap
            ElseIf NoOfAxles.Value = "Triple Axle" Then
               SWAxleClearance = 108                      ' Axle Wheel Clearance
               SWRailOverlap = 18                         ' Wall Rail Overlap
            End If
         ElseIf AxleSpacing.Value = "Spread Axle Spacing" Then
            If NoOfAxles.Value = "Dual Axle" Then
               SWAxleClearance = 82                       ' Axle Wheel Clearance
               SWRailOverlap = 18                         ' Wall Rail Overlap
            ElseIf NoOfAxles.Value = "Triple Axle" Then
               SWAxleClearance = 128                      ' Axle Wheel Clearance
               SWRailOverlap = 18                         ' Wall Rail Overlap
            End If
         End If
      End If
      AxleCL = TrailerLength.Value * AxleLocation.Value
      
End Function

Function TubeSize()
   
   If TrailerType.Value = "Motiv Raven Lite Cargo (Steel)" Then ' Raven Only
      If TrailerWidth.Value < 84 Then
         MainTubeHeight = 3
      Else
         MainTubeHeight = 4
      End If
   Else
      If TrailerLength.Value < 240 Then  ' ************************************Determine Frame Main Tube Sizes *****************************************
         If NoOfAxles.Value = "Dual Axle" Then
            If AxleRating.Value = "3500 lb 0 Deg" Then
               If TrailerWidth.Value < 84 Then
                   MainTubeHeight = 4
               Else
                  MainTubeHeight = 5
               End If
            ElseIf AxleRating.Value = "5200 lb 10 Deg up" Then
               If TrailerWidth.Value < 84 Then
                  MainTubeHeight = 4
               Else
                  MainTubeHeight = 5
               End If
            ElseIf AxleRating.Value = "6000 lb 10 Deg up" Then
               MainTubeHeight = 5
            ElseIf AxleRating.Value = "6000 lb 22.5 Deg up" Then
               MainTubeHeight = 5
            ElseIf AxleRating.Value = "7000 lb 10 Deg up" Then
               MainTubeHeight = 5
            ElseIf AxleRating.Value = "7000 lb 22.5 Deg up" Then
               MainTubeHeight = 5
            Else
               MainTubeHeight = 5
            End If
         ElseIf NoOfAxles.Value = "Triple Axle" Then
            If AxleRating.Value = "3500 lb 0 Deg" Then
               If TrailerWidth.Value < 84 Then
                  MainTubeHeight = 4
               Else
                  MainTubeHeight = 5
               End If
            Else
               MainTubeHeight = 5
            End If
         End If
      ElseIf TrailerLength.Value >= 240 And TrailerLength.Value < 288 Then
         MainTubeHeight = 5
      ElseIf TrailerLength.Value >= 288 And TrailerLength.Value < 336 Then
         If AxleRating.Value = "3500 lb 0 Deg" Or AxleRating.Value = "5200 lb 10 Deg up" Then
               MainTubeHeight = 5
            ElseIf AxleRating.Value = "6000 lb 10 Deg up" Then
               MainTubeHeight = 5
            ElseIf AxleRating.Value = "6000 lb 22.5 Deg up" Then
               MainTubeHeight = 5
            ElseIf AxleRating.Value = "7000 lb 10 Deg up" Then
               MainTubeHeight = 5
            ElseIf AxleRating.Value = "7000 lb 22.5 Deg up" Then
               MainTubeHeight = 5
            Else
               MainTubeHeight = 8
            End If
         ElseIf NoOfAxles.Value = "Triple Axle" Then
            If AxleRating.Value = "3500 lb 0 Deg" Then
               MainTubeHeight = 5
            Else
               MainTubeHeight = 8
            End If
       ElseIf TrailerLength.Value >= 336 Then
         If AxleRating.Value = "3500 lb 0 Deg" Or AxleRating.Value = "5200 lb 10 Deg up" Then
               MainTubeHeight = 8
         Else
               MainTubeHeight = 8
         End If
      End If ' ************************************Determine Frame Main Tube Sizes *****************************************
   End If ' End of Raven
   
   If MainTubeHeight = 3 Then                        ' *************** Set Frame Member Sizes ******************************
      FrameMember.Value = "Tube 2.0 x 3.0 x 11ga"
      InsideChannel = "C-CHANNEL 1.5 x 2.75 x 13ga"
      OverChannel = "C-CHANNEL 1.5 x 3.38 x 13ga"
      TGL = 5 ' Tongue Gusset Length
   ElseIf MainTubeHeight = 4 Then
      FrameMember.Value = "Tube 2.0 x 4.0 x 11ga"
      InsideChannel = "C-CHANNEL 1.5 x 3.75 x 13ga"
      OverChannel = "C-CHANNEL 1.5 x 4.38 x 13ga"
      TGL = 7 ' Tongue Gusset Length
   ElseIf MainTubeHeight = 5 Then
      FrameMember.Value = "Tube 2.0 x 5.0 x 11ga"
      InsideChannel = "C-CHANNEL 1.5 x 4.88 x 13ga"
      OverChannel = "C-CHANNEL 1.5 x 5.38 x 13ga"
      TGL = 7.5 ' Tongue Gusset Length
   ElseIf MainTubeHeight = 8 Then
      FrameMember.Value = "Tube 2.0 x 8.0 x 11ga"
      InsideChannel = "C-CHANNEL 1.5 x 4.88 x 13ga"
      OverChannel = "C-CHANNEL 1.5 x 8.38 x 13ga"
      TGL = 7.5 ' Tongue Gusset Length
   End If
   
   
      
End Function

Function TongueDetails()
   
   If TrailerType.Value = "Motiv Raven Lite Cargo (Steel)" Then 'Determine Tongue Length **************Separate into New Function *******************
      If NoseStructure.Value = "Flat Nose" Then
         If TrailerWidth.Value = "60" Then                       ' 5 foot trailer
            TongueLength.Value = 39
         ElseIf TrailerWidth.Value = "72" Then                   ' 6 foot trailer
            TongueLength.Value = 39
         ElseIf TrailerWidth.Value = "84" Then                   ' 7 foot trailer
            TongueLength.Value = 39
         End If
      ElseIf NoseStructure.Value = "2ft Wedge" Then
         If TrailerWidth.Value = "60" Then                       ' 5 foot trailer
            TongueLength.Value = 42
         ElseIf TrailerWidth.Value = "72" Then                   ' 6 foot trailer
            TongueLength.Value = 47
         ElseIf TrailerWidth.Value = "84" Then                   ' 7 foot trailer
            TongueLength.Value = 47
         End If
      End If
   End If

End Function

Function CreateConfigurations()

    Dim swApp                       As SldWorks.SldWorks
    Dim swModel                     As SldWorks.ModelDoc2
    Dim swAssy                      As SldWorks.AssemblyDoc
    Dim swConfMgr                   As SldWorks.ConfigurationManager
    Dim swConf                      As SldWorks.Configuration

    Set swApp = CreateObject("SldWorks.Application")
    Set swModel = swApp.ActiveDoc
    Set swConfMgr = swModel.ConfigurationManager
    Set swConf = swConfMgr.ActiveConfiguration
   
   'Dim swApp As Object

   Dim Part As Object
   Dim boolstatus As Boolean
   Dim longstatus As Long, longwarnings As Long
   
   Set swApp = Application.SldWorks
   Set Part = swApp.ActiveDoc
   'Set CongifM = SldWorks.ConfigurationManager

   boolstatus = Part.Extension.SelectByID2(CStr(PartNumber.Value) + ".SLDPRT", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
   boolstatus = Part.Extension.SelectByID2(CStr(PartNumber.Value) + ".SLDPRT", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
   Part.ClearSelection2 True
   boolstatus = Part.AddConfiguration2("Frame", "", "", True, False, False, True, 256)
   boolstatus = Part.Extension.SelectByID2(CStr(PartNumber.Value) + ".SLDPRT", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
   Part.ClearSelection2 True
   boolstatus = Part.AddConfiguration2("Front Wall", "", "", True, False, False, True, 256)
   'boolstatus = swConfMgr.SetConfigurationParams("Front Wall", "Description", "Front Wall")
   Part.ClearSelection2 True
   boolstatus = Part.AddConfiguration2("CS Wall", "", "", True, False, False, True, 256)
   Part.ClearSelection2 True
   boolstatus = Part.AddConfiguration2("RS Wall", "", "", True, False, False, True, 256)
   Part.ClearSelection2 True
   boolstatus = Part.AddConfiguration2("Rear Wall", "", "", True, False, False, True, 256)
   Part.ClearSelection2 True
   boolstatus = Part.AddConfiguration2("Roof", "", "", True, False, False, True, 256)
   
End Function

Function CreateConfigurations2()

   Dim swApp               As SldWorks.SldWorks
   Dim swModel             As SldWorks.ModelDoc2
   Dim swAssy              As SldWorks.AssemblyDoc

   Dim iswConfigMgr        As SldWorks.IConfigurationManager
   Dim swConfigMgr         As SldWorks.ConfigurationManager
   Dim swConfMgr           As SldWorks.ConfigurationManager
   Dim swConfig            As SldWorks.Configuration
   Dim swConf              As SldWorks.Configuration
   Dim swCustPropMgr       As SldWorks.CustomPropertyManager
    
   Dim vConfName           As Variant
   Dim vPropName           As Variant
   Dim vPropNames          As Variant
   Dim vPropValue          As Variant
   Dim vPropType           As Variant
   Dim nNumProp            As Long
   Dim i                   As Long
   Dim j                   As Long
   Dim bRet                As Boolean
   Dim ConfigName          As String
    
   Dim Part                As Object
   Dim boolstatus          As Boolean
   Dim longstatus          As Long, longwarnings As Long

   Dim nNbrProps           As Long
   Dim retVal              As Long
   Dim valOut              As String
   Dim resolvedValOut      As String
   Dim custPropType        As Long
   
   Dim TLen                As Double
   Dim NS                  As String
   Dim CSDW                As String
   Dim RSDW                As String
   Dim FRM                 As String
   Dim ROOF                As String
   Dim FRW                 As String
   Dim RRW                 As String
   Dim RAMP                As String
   
   'Dim instance            As IConfigurationManager
   ' ICustomPropertyManager
   ' IModelDocExtension
 

   Set swApp = CreateObject("SldWorks.Application")
   Set swModel = swApp.ActiveDoc
   Set swConfMgr = swModel.ConfigurationManager
   Set swConf = swConfMgr.ActiveConfiguration
   Set swApp = Application.SldWorks
   Set Part = swApp.ActiveDoc
   
   Set swApp = Application.SldWorks
   Set swModel = swApp.ActiveDoc
   
   Set swApp = Application.SldWorks
   Set swModel = swApp.ActiveDoc
   
   Set swConfMgr = swModel.ConfigurationManager
   Set swConf = swConfMgr.ActiveConfiguration
   
   ' Configuration Names
   
   'SDW, ST - 10X72+0SA
   If NoseStructure.Value = "Flat Nose" Then
      NS = "+0"
   ElseIf NoseStructure.Value = "2ft Wedge" Then
      NS = "+2"
   ElseIf NoseStructure.Value = "4ft Wedge" Then
      NS = "+4"
   ElseIf NoseStructure.Value = "6ft Wedge" Then
      NS = "+6"
   End If
   TLen = (TrailerLength.Value - 2) / 12
   
   CSDW = "SDW-CS, ST - " + CStr(TLen) + "X" + CStr(TrailerHeight.Value) + NS + "SA"
   RSDW = "SDW-RS, ST - " + CStr(TLen) + "X" + CStr(TrailerHeight.Value) + NS + "SA"
   
   FRM = "FRM, ST - " + CStr(TLen) + "X" + CStr(TrailerHeight.Value) + NS + "SA"
   ROOF = "ROOF, ST - " + CStr(TLen) + "X" + CStr(TrailerHeight.Value) + NS + "SA"
   FRW = "FRW, ST - " + CStr(TLen) + "X" + CStr(TrailerHeight.Value) + NS + "SA"
   RRW = "RRW, ST - " + CStr(TLen) + "X" + CStr(TrailerHeight.Value) + NS + "SA"
   RAMP = "RAMP, ST - " + CStr(TLen) + "X" + CStr(TrailerHeight.Value) + NS + "SA"
   

   boolstatus = Part.Extension.SelectByID2(CStr(PartNumber.Value) + ".SLDPRT", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
   boolstatus = Part.Extension.SelectByID2(CStr(PartNumber.Value) + ".SLDPRT", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
   Part.ClearSelection2 True
   'boolstatus = Part.AddConfiguration2("Frame", "", "", True, False, False, True, 256)
   boolstatus = Part.AddConfiguration2(FRM, "", "", True, False, False, True, 256)
   
   vConfName = swModel.GetConfigurationNames
   Set swConfig = swModel.GetConfigurationByName(vConfName(UBound(vConfName)))
   nNumProp = swConfig.GetCustomProperties(vPropName, vPropValue, vPropType)
   ConfigName = vConfName(UBound(vConfName))
   Set swCustPropMgr = swConfig.CustomPropertyManager                       ' Get the number of custom properties for this configuration
   retVal = swCustPropMgr.Add2("Date ", swCustomInfoDate, "24-Feb-2011")    ' Get the new number of custom properties for this configuration
   retVal = swCustPropMgr.Add2("Description ", 30, "Frame")                 ' Get the new number of custom properties for this configuration
   retVal = swCustPropMgr.Add2("Part No ", 30, "123456")                    ' Get the new number of custom properties for this configuration
   
   
   'boolstatus = Part.Extension.SelectByID2(CStr(PartNumber.Value) + ".SLDPRT", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
   Part.ClearSelection2 True
   'boolstatus = Part.AddConfiguration2("Front Wall", "", "", True, False, False, True, 256)
   boolstatus = Part.AddConfiguration2(FRW, "", "", True, False, False, True, 256)
   Part.ClearSelection2 True
   'boolstatus = Part.AddConfiguration2("CS Wall", "", "", True, False, False, True, 256)
   boolstatus = Part.AddConfiguration2(CSDW, "", "", True, False, False, True, 256)
   Part.ClearSelection2 True
   'boolstatus = Part.AddConfiguration2("RS Wall", "", "", True, False, False, True, 256)
   boolstatus = Part.AddConfiguration2(RSDW, "", "", True, False, False, True, 256)
   Part.ClearSelection2 True
   'boolstatus = Part.AddConfiguration2("Rear Wall", "", "", True, False, False, True, 256)
   boolstatus = Part.AddConfiguration2(RRW, "", "", True, False, False, True, 256)
   Part.ClearSelection2 True
   'boolstatus = Part.AddConfiguration2("Roof", "", "", True, False, False, True, 256)
   boolstatus = Part.AddConfiguration2(ROOF, "", "", True, False, False, True, 256)
   Part.ClearSelection2 True
   'boolstatus = Part.AddConfiguration2("Ramp", "", "", True, False, False, True, 256)
   boolstatus = Part.AddConfiguration2(RAMP, "", "", True, False, False, True, 256)
    
   'vConfName = swModel.GetConfigurationNames
   
   'Set swConfig = swModel.GetConfigurationByName(vConfName(UBound(vConfName)))
   'nNumProp = swConfig.GetCustomProperties(vPropName, vPropValue, vPropType)
   'ConfigName = vConfName(UBound(vConfName))
       
  ' ICustomPropertyManager
  ' IModelDocExtension
      
   'Set swCustPropMgr = swConfig.CustomPropertyManager    ' Get the number of custom properties for this configuration
    
   'retVal = swCustPropMgr.Add2("Date ", swCustomInfoDate, "24-Feb-2011")    ' Get the new number of custom properties for this configuration
   'retVal = swCustPropMgr.Add2("Description ", 30, "Frame")    ' Get the new number of custom properties for this configuration
   'retVal = swCustPropMgr.Add2("Part No ", 30, "123456")    ' Get the new number of custom properties for this configuration
   
End Function

Function BOMPartNumber(config As SldWorks.Configuration, document As SldWorks.ModelDoc2) As String

    Select Case config.BOMPartNoSource
    Case SwConst.swBOMPartNumberSource_e.swBOMPartNumber_ConfigurationName
        BOMPartNumber = config.Name
    Case SwConst.swBOMPartNumberSource_e.swBOMPartNumber_DocumentName
        BOMPartNumber = document.GetTitle
    Case SwConst.swBOMPartNumberSource_e.swBOMPartNumber_UserSpecified
        BOMPartNumber = config.AlternateName
    Case SwConst.swBOMPartNumberSource_e.swBOMPartNumber_ParentName
        Dim parentConfig As SldWorks.Configuration
        Set parentConfig = config.GetParent
        If parentConfig.BOMPartNoSource = SwConst.swBOMPartNumberSource_e.swBOMPartNumber_ParentName Then
            BOMPartNumber = BOMPartNumber(parentConfig, document)
        Else
            BOMPartNumber = parentConfig.Name
        End If
    End Select
End Function

Function InspectConfigurations(Doc As SldWorks.ModelDoc2)

    Dim params As Variant
    params = Doc.GetConfigurationNames
    Dim vName As Variant
    Dim Name As String
    Dim thisConfig As Configuration
    For Each vName In params
        Name = vName
        Set thisConfig = Doc.GetConfigurationByName(Name)
        
        Debug.Print "Name                      ", thisConfig.Name
        
        ' Work out what the BOM part number is based on any derived configurations
        Debug.Print "BOMPartNumber             ", BOMPartNumber(thisConfig, Doc)
        Debug.Print "AlternateName             ", thisConfig.AlternateName
        Debug.Print "Comment                   ", thisConfig.Comment
        Debug.Print "Description               ", thisConfig.Description
        Debug.Print "HideNewComponentModels    ", thisConfig.HideNewComponentModels
        Debug.Print "Lock                      ", thisConfig.Lock
        Debug.Print "ShowChildComponentsInBOM  ", thisConfig.ShowChildComponentsInBOM
        Debug.Print "UseAlternateNameInBOM     ", thisConfig.UseAlternateNameInBOM
        Debug.Print "SuppressNewComponentModels", thisConfig.SuppressNewComponentModels
        Debug.Print "SuppressNewFeatures       ", thisConfig.SuppressNewFeatures
        Debug.Print "------------------------------------------------------------------"
    Next vName
End Function
 

Function yetagain()
 
Dim swApp  As SldWorks.SldWorks
Dim Part   As SldWorks.PartDoc
Dim Doc    As SldWorks.ModelDoc2
Dim SelMgr As SldWorks.SelectionMgr
 
Set swApp = Application.SldWorks
Set Part = swApp.ActiveDoc
Set Doc = Part
Set SelMgr = Doc.SelectionManager

Call InspectConfigurations(Doc)

Dim params As Variant
params = Doc.GetConfigurationNames
Dim vName As Variant
Dim Name As String
Dim thisConfig As Configuration
MsgBox ("Modifying the configurations...")
For Each vName In params
    Name = vName
    Set thisConfig = Doc.GetConfigurationByName(Name)
    MsgBox ("Name                   :" + thisConfig.Name)

    thisConfig.BOMPartNoSource = swBOMPartNumber_UserSpecified
    thisConfig.AlternateName = "XXXX"
    thisConfig.UseAlternateNameInBOM = True
    thisConfig.AlternateName = "XXXX"
    thisConfig.Description = "This Test"

Next vName
 
Call InspectConfigurations(Doc)

End Function

'Public Enum swCustomInfoType_e
'    swCustomInfoUnknown = 0
'    swCustomInfoText = 30       '  VT_LPSTR
'    swCustomInfoDate = 64       '  VT_FILETIME
'    swCustomInfoNumber = 3      '  VT_I4
'    swCustomInfoYesOrNo = 11    '  VT_BOOL
'End Enum
 
Function ConfigProp()

    Dim swApp                   As SldWorks.SldWorks
    Dim swModel                 As SldWorks.ModelDoc2

    Dim iswConfigMgr            As SldWorks.IConfigurationManager
    
    Dim swConfig                As SldWorks.Configuration
    Dim swAssy                  As SldWorks.AssemblyDoc
    Dim swConfMgr               As SldWorks.ConfigurationManager
    Dim swConf                  As SldWorks.Configuration
    Dim vConfName               As Variant
    Dim vPropName               As Variant
    Dim vPropValue              As Variant
    Dim vPropType               As Variant
    Dim nNumProp                As Long
    Dim i                       As Long
    Dim j                       As Long
    Dim bRet                    As Boolean
    Dim ConfigName As String
    
    Dim Part As Object
    Dim boolstatus As Boolean
    Dim longstatus As Long, longwarnings As Long
    
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    'Set swApp = CreateObject("SldWorks.Application")
    
    Set swConfMgr = swModel.ConfigurationManager
    Set swConf = swConfMgr.ActiveConfiguration

    'MsgBox ("File = " + swModel.GetPathName)

    vConfName = swModel.GetConfigurationNames
    For i = 0 To UBound(vConfName)
        Set swConfig = swModel.GetConfigurationByName(vConfName(i))
        nNumProp = swConfig.GetCustomProperties(vPropName, vPropValue, vPropType)
        ConfigName = vConfName(i)
        
        MsgBox ("Config       = " & vConfName(i))
        
        For j = 0 To nNumProp - 1
            boolstatus = swConfMgr.SetConfigurationParams(ConfigName, vPropName(j), "TEST")
            'MsgBox ("    " & vPropName(j) & " <" & vPropType(j) & "> = " & vPropValue(j))
        Next j
    Next i
  
  ' ICustomPropertyManager
  ' IModelDocExtension
  
    'Dim swApp               As SldWorks.SldWorks
    'Dim swModel             As SldWorks.ModelDoc2
    Dim swConfigMgr         As SldWorks.ConfigurationManager
    'Dim swConfig            As SldWorks.Configuration
    Dim swCustPropMgr       As SldWorks.CustomPropertyManager
    Dim nNbrProps           As Long
    'Dim j                   As Long
    Dim retVal              As Long
    Dim vPropNames          As Variant
    Dim valOut              As String
    Dim resolvedValOut      As String
    Dim custPropType        As Long
  
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    'Set swConfigMgr = swModel.ConfigurationManager
    'Set swConfig = swConfigMgr.ActiveConfiguration
    MsgBox ("Name of this configuration:                             " & swConfig.Name)
  
    Set swCustPropMgr = swConfig.CustomPropertyManager    ' Get the number of custom properties for this configuration
    nNbrProps = swCustPropMgr.Count
    MsgBox ("Number of properties for this configuration:            " & nNbrProps)    ' Add custom property date to this configuration
    
    retVal = swCustPropMgr.Add2("Date ", swCustomInfoDate, "24-Feb-2011")    ' Get the new number of custom properties for this configuration
    retVal = swCustPropMgr.Add2("Description ", 30, "Frame")    ' Get the new number of custom properties for this configuration
    nNbrProps = swCustPropMgr.Count
    
    MsgBox ("New number of properties for this configuration:        " & nNbrProps)    ' Get the names of the custom properties
    vPropNames = swCustPropMgr.GetNames
    
   ' For each custom property, get its type, value, and resolved value
   ' Then print its name, type, and resolved value
    For j = 0 To nNbrProps - 1
        swCustPropMgr.Get2 vPropNames(j), valOut, resolvedValOut
        custPropType = swCustPropMgr.GetType2(vPropNames(j))
        MsgBox ("    Name, type, and resolved value of custom property:  " & vPropNames(j) & " - " & custPropType & " - " & resolvedValOut)
    Next j
  
End Function

Function CreateDerivedConfig()
    
    Dim swApp                       As SldWorks.SldWorks
    Dim swModel                     As SldWorks.ModelDoc2
    Dim vConfigNameArr              As Variant
    Dim vConfigName                 As Variant
    Dim vConfigParam                As Object
    Dim vParam                      As Object
    Dim vValues                     As Object
    Dim swActiveConf                As SldWorks.Configuration
    Dim swConf                      As SldWorks.Configuration
    Dim swConfMgr                   As SldWorks.ConfigurationManager
    Dim swDerivConf                 As SldWorks.Configuration
    Dim bRet                        As Boolean
    
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set swConfMgr = swModel.ConfigurationManager
    'Set swConfig = swConfMgr.ActiveConfiguration

    vConfigNameArr = swModel.GetConfigurationNames

    For Each vConfigName In vConfigNameArr
        Set swConf = swModel.GetConfigurationByName(vConfigName)
        
        MsgBox (swConf.Name)
        
        'Set vConfigParam = swConfMgr.GetConfigurationParams(swConf.Name, vParam, vValues)
                
        'vConfigParam.
        
        'MsgBox (vConfigParam.Name)
        
        ' Do not assert; will be NULL if (derived) configuration already exists

        Set swDerivConf = swConfMgr.AddConfiguration(swConf.Name + " Derived", "Derived comment", "Derived alternate name", 0, swConf.Name, "Derived description")
    Next

End Function

Function CreatePlanes()

Dim OCD As Double
Dim TW As Double
Dim Weldment As Double
Dim LineCnt As Integer
Dim X As Integer
Dim i As Integer

'110207
Dim FramePlane_Offset As Double
Dim FWPlane_Offset As Double
Dim FWR_Offset(10) As Double
Dim RWPlane_Offset As Double
Dim CSPlane_Offset As Double
Dim RSPlane_Offset As Double
Dim RoofPlane_Offset As Double

Dim swApp As Object
Dim Part As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long
Dim cnt As Integer

Set swApp = _
Application.SldWorks

Set Part = swApp.ActiveDoc

' Set Object Variable(s)
Set swApp = Application.SldWorks
Set Part = swApp.ActiveDoc

' Setting the Input Parameters

TW = 2

OCD = OCDist.Value

FramePlane_Offset = ((MainTubeHeight / 2) * 0.0254)                    ' Frame Offset
FWPlane_Offset = 0.0127                                                ' Front Wall Plane Offset, needs to be based on Wall Thickness
RWPlane_Offset = (0.0254 * TrailerLength.Value)                        ' Rear Wall Offset
CSPlane_Offset = ((0.0254 * TrailerWidth.Value) / 2) - 0.0127          ' Curb Side Wall Offset, needs to be based on Wall Thickness
RSPlane_Offset = ((0.0254 * TrailerWidth.Value) / 2) - 0.0127          ' Road Side Wall Offset, needs to be based on Wall Thickness
RoofPlane_Offset = (TrailerTotalHeight - SWHeaderTubeHeight / 2) * 0.0254 ' Roof Height Adjusted for RoofTubeHeight

If NoseStructure.Value = "2ft Wedge" Then
   If TrailerTotalHeight <= 72 Then
      FWR_Offset(1) = 0.0127                                               ' Bottom Member - Offset Frame
      FWR_Offset(2) = 0.6096 + 0.0127                                      ' @ 24 Inches for Gravel Guard
      FWR_Offset(3) = (((TrailerTotalHeight - 24) / 2) + 24) * 0.0254     ' Split distance between Upper and Middle
      FWR_Offset(4) = (TrailerTotalHeight * 0.0254) - 0.0127              ' Top Member - Offset Frame
      X = 4
   Else
      FWR_Offset(1) = 0.0127                                               ' Bottom Member - Offset Frame
      FWR_Offset(2) = 0.6096 + 0.0127                                      ' @ 24 Inches for Gravel Guard
      FWR_Offset(3) = (((TrailerTotalHeight - 24) / 3) + 24) * 0.0254     ' Split distance between Upper and Middle
      FWR_Offset(4) = (((TrailerTotalHeight - 24) / 3) * 2 + 24) * 0.0254 ' Split distance between Upper and Middle
      FWR_Offset(5) = (TrailerTotalHeight * 0.0254) - 0.0127              ' Top Member - Offset Frame
      X = 5
   End If
End If

i = 1

' Road Side Wall
boolstatus = Part.Extension.SelectByID2("Right", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
Part.CreatePlaneAtOffset3 RSPlane_Offset, False, True
Part.ClearSelection2 True
boolstatus = Part.Extension.SelectByID2("Plane" + CStr(i), "PLANE", 0, 0, 0, False, 0, Nothing, 0)
Part.SelectedFeatureProperties 0, 0, 0, 0, 0, 0, 0, 1, 0, "RS"
i = i + 1

' Curb Side Wall
boolstatus = Part.Extension.SelectByID2("Right", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
Part.CreatePlaneAtOffset3 CSPlane_Offset, True, True
Part.ClearSelection2 True
boolstatus = Part.Extension.SelectByID2("Plane" + CStr(i), "PLANE", 0, 0, 0, False, 0, Nothing, 0)
Part.SelectedFeatureProperties 0, 0, 0, 0, 0, 0, 0, 1, 0, "CS"
i = i + 1

' Front Wall
boolstatus = Part.Extension.SelectByID2("Front", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
Part.CreatePlaneAtOffset3 FWPlane_Offset, True, True
Part.ClearSelection2 True
boolstatus = Part.Extension.SelectByID2("Plane" + CStr(i), "PLANE", 0, 0, 0, False, 0, Nothing, 0)
Part.SelectedFeatureProperties 0, 0, 0, 0, 0, 0, 0, 1, 0, "FW"
i = i + 1

' Rear Wall
boolstatus = Part.Extension.SelectByID2("Front", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
Part.CreatePlaneAtOffset3 RWPlane_Offset, True, True
Part.ClearSelection2 True
boolstatus = Part.Extension.SelectByID2("Plane" + CStr(i), "PLANE", 0, 0, 0, False, 0, Nothing, 0)
Part.SelectedFeatureProperties 0, 0, 0, 0, 0, 0, 0, 1, 0, "RW"
i = i + 1

' Frame
boolstatus = Part.Extension.SelectByID2("Top", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
Part.CreatePlaneAtOffset3 FramePlane_Offset, True, True
Part.ClearSelection2 True
boolstatus = Part.Extension.SelectByID2("Plane" + CStr(i), "PLANE", 0, 0, 0, False, 0, Nothing, 0)
Part.SelectedFeatureProperties 0, 0, 0, 0, 0, 0, 0, 1, 0, "FRAME"
i = i + 1

For cnt = 1 To X
' Frame (Raven Wedge Nose)
   boolstatus = Part.Extension.SelectByID2("Top", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
   Part.CreatePlaneAtOffset3 FWR_Offset(cnt), False, True
   Part.ClearSelection2 True
   boolstatus = Part.Extension.SelectByID2("Plane" + CStr(i), "PLANE", 0, 0, 0, False, 0, Nothing, 0)
   Part.SelectedFeatureProperties 0, 0, 0, 0, 0, 0, 0, 1, 0, "FWR" + CStr(cnt)
   i = i + 1
Next cnt

' Roof
boolstatus = Part.Extension.SelectByID2("Top", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
Part.CreatePlaneAtOffset3 RoofPlane_Offset, False, True
Part.ClearSelection2 True
boolstatus = Part.Extension.SelectByID2("Plane" + CStr(i), "PLANE", 0, 0, 0, False, 0, Nothing, 0)
Part.SelectedFeatureProperties 0, 0, 0, 0, 0, 0, 0, 1, 0, "ROOF"
i = i + 1

' Create Axis Along Z
boolstatus = Part.Extension.SelectByID2("Front", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
boolstatus = Part.Extension.SelectByID2("Right", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
boolstatus = Part.InsertAxis2(True)
boolstatus = Part.Extension.SelectByID2("Axis1", "AXIS", 0, 0, 0, False, 0, Nothing, 0)
boolstatus = Part.Extension.SelectByID2("Axis1", "AXIS", 0, 0, 0, False, 0, Nothing, 0)
boolstatus = Part.SelectedFeatureProperties(0, 0, 0, 0, 0, 0, 0, 1, 0, "VerticalAxis")



   Dim PlaneRot As Double
   Dim PlaneOffset As Double
  
   PlaneRot = Atn(12 / ((TrailerWidth.Value / 2) - 12))
   PlaneOffset = Sin(PlaneRot) * ((TrailerWidth.Value / 2) - 12) + 12
  
   'CSPlane_Offset = (21.984609) * 0.0254                                 ' Curb Side Offset Plane for Raven Wedge Backers
   'RSPlane_Offset = (21.984609) * 0.0254                                 ' Road Side Offset Plane for Raven Wedge Backers

   CSPlane_Offset = (PlaneOffset) * 0.0254                                 ' Curb Side Offset Plane for Raven Wedge Backers
   RSPlane_Offset = (PlaneOffset) * 0.0254                                 ' Road Side Offset Plane for Raven Wedge Backers

   ' Plane for Raven Wedge RS Plane *** For Backer ***
   Part.ClearSelection2 True
   boolstatus = Part.Extension.SelectByID2("VerticalAxis", "AXIS", 0, 0, 0, True, 0, Nothing, 0)
   boolstatus = Part.Extension.SelectByID2("Front", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
   Dim myRefPlane As Object
   Set myRefPlane = Part.FeatureManager.InsertRefPlane(4, 0, 16, PlaneRot, 0, 0)
   Part.ClearSelection2 True
   boolstatus = Part.Extension.SelectByID2("Plane" + CStr(i), "PLANE", 0, 0, 0, False, 0, Nothing, 0)
   Part.SelectedFeatureProperties 0, 0, 0, 0, 0, 0, 0, 1, 0, "CSB"
   i = i + 1

   ' Plane for Raven Wedge CS Plane *** For Backer ***
   Part.ClearSelection2 True
   boolstatus = Part.Extension.SelectByID2("VerticalAxis", "AXIS", 0, 0, 0, True, 0, Nothing, 0)
   boolstatus = Part.Extension.SelectByID2("Front", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
   Set myRefPlane = Part.FeatureManager.InsertRefPlane(4, 0, 272, PlaneRot, 0, 0)
   Part.ClearSelection2 True
   boolstatus = Part.Extension.SelectByID2("Plane" + CStr(i), "PLANE", 0, 0, 0, False, 0, Nothing, 0)
   Part.SelectedFeatureProperties 0, 0, 0, 0, 0, 0, 0, 1, 0, "RSB"
   i = i + 1

' *** Road Side Offset Plane for Raven Wedge Backers ***
Part.ClearSelection2 True
boolstatus = Part.Extension.SelectByID2("RSB", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
Part.CreatePlaneAtOffset3 RSPlane_Offset, False, True
Part.ClearSelection2 True
boolstatus = Part.Extension.SelectByID2("Plane" + CStr(i), "PLANE", 0, 0, 0, False, 0, Nothing, 0)
Part.SelectedFeatureProperties 0, 0, 0, 0, 0, 0, 0, 1, 0, "RSBP"
i = i + 1

' *** Curb Side Offset Plane for Raven Wedge Backers ***
Part.ClearSelection2 True
boolstatus = Part.Extension.SelectByID2("CSB", "PLANE", 0, 0, 0, True, 0, Nothing, 0)
Part.CreatePlaneAtOffset3 CSPlane_Offset, False, True
Part.ClearSelection2 True
boolstatus = Part.Extension.SelectByID2("Plane" + CStr(i), "PLANE", 0, 0, 0, False, 0, Nothing, 0)
Part.SelectedFeatureProperties 0, 0, 0, 0, 0, 0, 0, 1, 0, "CSBP"
i = i + 1

Part.ClearSelection2 True

End Function

Private Sub CB_Exit_Click()
Dim swApp As Object

Dim Part As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long

Set swApp = _
Application.SldWorks

Set Part = swApp.ActiveDoc

boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swDisplayPlanes, False)
boolstatus = Part.SetUserPreferenceToggle(swUserPreferenceToggle_e.swDisplaySketches, False)

swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchInference, True

'Set Part = swApp.ActiveDoc
'boolstatus = Part.Extension.SelectByID2("Top", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
'Part.SketchManager.InsertSketch True
'Part.ClearSelection2 True
'Dim skSegment As Object
'Set skSegment = Part.SketchManager.CreateLine(3.332042, -0.408209, 0#, 2.486985, 0.807486, 0#)
'Part.SetPickMode
'Part.ClearSelection2 True
'boolstatus = Part.EditRebuild3()
'Dim myFeature As Object
'Set myFeature = Part.FeatureManager.InsertWeldmentFeature()
'boolstatus = Part.Extension.SelectByID2("Line1@Sketch117", "EXTSKETCHSEGMENT", 3.00615302876, 0.06061327381209, 0, True, 0, Nothing, 0)
'Dim vGroups As Variant
'Dim GroupArray() As Object
'ReDim GroupArray(0 To 0) As Object
'Dim Group1 As Object
'Set Group1 = Part.FeatureManager.CreateStructuralMemberGroup()
'Dim vSegement1 As Variant
'Dim SegementArray1() As Object
'ReDim SegementArray1(0 To 0) As Object
'Part.ClearSelection2 True
'boolstatus = Part.Extension.SelectByID2("Line1@Sketch117", "EXTSKETCHSEGMENT", 5.392792444839, 0, 1.349630948817, True, 0, Nothing, 0)
'Dim Segment As Object
'Set Segment = Part.SelectionManager.GetSelectedObject5(1)
'Set SegementArray1(0) = Segment
'vSegement1 = SegementArray1
'Group1.Segments = (vSegement1)
'Group1.ApplyCornerTreatment = True
'Group1.CornerTreatmentType = 1
'Group1.GapWithinGroup = 0
'Group1.GapForOtherGroups = 0
'Group1.Angle = 0
'Set GroupArray(0) = Group1
'vGroups = GroupArray
'Set myFeature = Part.FeatureManager.InsertStructuralWeldment4("C:\ATC ENGINEERING\EXTRUSIONS\STEEL\TUBES\2 x 8 x 11ga.sldlfp", 1, True, (vGroups))
'Part.ClearSelection2 True
'boolstatus = Part.Extension.SelectByID2("Structural Member18", "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
'Part.EditDelete
'boolstatus = Part.Extension.SelectByID2("Sketch117", "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
'Part.EditDelete

End

End Sub


Private Sub CB10_Click()

   Dim swApp As Object
   Dim Part As Object
   Dim SelMgr As Object
   Dim boolstatus As Boolean
   Dim longstatus As Long, longwarnings As Long
   Dim Feature As Object

   Set swApp = Application.SldWorks

   Set Part = swApp.ActiveDoc
   Set SelMgr = Part.SelectionManager
   boolstatus = Part.EditRebuild3

End Sub

Private Sub CB11_Click()


End Sub

Private Sub CB12_Click()

End Sub

Private Sub CB13_Click()
 
   Dim TW As Double
   Dim Weldment As Double
   Dim LineCnt As Integer
   Dim X As Integer
   Dim i As Integer
   ' 110207 *************************************************************************************
   Dim WOCD As Double
   Dim Y As Double
   Dim WeldmentType(100) As String
   Dim SketchName As String

    ' Add Variable for Sketch Plane
   Dim SketchPlane As String
   '**********************************************************************************************

   Dim swApp As Object
   Dim Part As Object
   Dim boolstatus As Boolean
   Dim longstatus As Long, longwarnings As Long
   Dim RotationAngle(100) As Double          ' Angle of Rotation of the Weldment Profile

   Set swApp = Application.SldWorks
   Set Part = swApp.ActiveDoc

   '# Setting the Input Parameters
   TW = 1
   LineCnt = 0
   WOCD = 24

   boolstatus = Part.Extension.SelectByID2("Front Wall@" + CStr(PartNumber.Value) + ".SLDPRT", "CONFIGURATIONS", 0, 0, 0, False, 0, Nothing, 0)
   boolstatus = Part.ShowConfiguration2("Front Wall")
   
   '110207 *************************************************************************************
   Part.ClearSelection2 True
   Dim skSegment As Object
   ' *******************************************************************************************

   If NoseStructure.Value = "Flat Nose" Then ' ******************************* Nose Structure: Flat *******************************

      ' Select Working Plane
      boolstatus = Part.Extension.SelectByID2("FW", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
      Part.SketchManager.InsertSketch True
      
      '# Wall Header **********************************************************
      P1x = -((TrailerWidth.Value / 2) - SW_Offset)
      P1y = TrailerTotalHeight - (SWHeaderTubeHeight / 2)
      P2x = ((TrailerWidth.Value / 2) - SW_Offset)
      P2y = TrailerTotalHeight - (SWHeaderTubeHeight / 2)
      ' Convert to Meters
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"

      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", "FRONT", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.

      '# Wall Footer ***********************************************************
      P1x = -((TrailerWidth.Value / 2) - SW_Offset)
      P1y = SWHeaderTubeHeight / 2
      P2x = ((TrailerWidth.Value / 2) - SW_Offset)
      P2y = SWHeaderTubeHeight / 2
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", "FRONT", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.

      ' CS Vertical Tube ********************************************************
   
      P1x = -((TrailerWidth.Value / 2) - SW_Offset - (SWFrontVerticalTubeHeight / 2))
      P1y = (SWHeaderTubeHeight)
      P2x = -((TrailerWidth.Value / 2) - SW_Offset - (SWFrontVerticalTubeHeight / 2))
      P2y = TrailerTotalHeight - (SWHeaderTubeHeight)
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Y = (i * WOCD)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"

      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", "FRONT", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.

      ' RS Vertical Tube *********************************************************

      P1x = ((TrailerWidth.Value / 2) - SW_Offset - (SWFrontVerticalTubeHeight / 2))
      P1y = (SWHeaderTubeHeight)
      P2x = ((TrailerWidth.Value / 2) - SW_Offset - (SWFrontVerticalTubeHeight / 2))
      P2y = TrailerTotalHeight - (SWHeaderTubeHeight)
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Y = (i * WOCD)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"

      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", "FRONT", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
      
      ' Raven Front Only
      ' Center Z Channel **********************************************************

      P1x = 0
      P1y = (SWHeaderTubeHeight)
      P2x = 0
      P2y = TrailerTotalHeight - (SWHeaderTubeHeight)
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Y = (i * WOCD)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = "Z-CHANNEL 1.125 x 1.125"

      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", "FRONT", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.

      '# Note: Repeat this Operation based on Length and Frame Center to Center Distance

      ' Z Channel ******************************************************************

      X = ((((TrailerWidth.Value / 2) - SW_Offset - 2)) - ((((TrailerWidth.Value / 2) - SW_Offset) - 2) Mod WOCD)) / WOCD
      WOCD = ((TrailerWidth.Value / 2) - SW_Offset) / (X + 1)
      
      For i = 1 To X
         P1x = WOCD * i
         P1y = (SWHeaderTubeHeight)
         P2x = WOCD * i
         P2y = TrailerTotalHeight - (SWHeaderTubeHeight)
         ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
         Y = (i * WOCD)
         Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
         LineCnt = LineCnt + 1
         'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
         RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
         WeldmentType(LineCnt) = "Z-CHANNEL 1.125 x 1.125"
         
         Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
         Write #1, "WALL", "FRONT", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
         
      Next i

      For i = 1 To X
         P1x = -WOCD * i
         P1y = (SWHeaderTubeHeight)
         P2x = -WOCD * i
         P2y = TrailerTotalHeight - (SWHeaderTubeHeight)
         ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
         Y = (i * WOCD)
         Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
         LineCnt = LineCnt + 1
         'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
         RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
         WeldmentType(LineCnt) = "Z-CHANNEL 1.125 x 1.125"
         
         Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
         Write #1, "WALL", "FRONT", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
         
      Next i
      
      ' *********************************************** Front Corner Pieces ***************************************************
      
      If TrailerWidth.Value = 84 Then
         P1y = -4
      Else
         P1y = -3
      End If
            
      P1x = TrailerWidth.Value / 2
      P1y = P1y
      P2x = TrailerWidth.Value / 2
      P2y = P1y + 6
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Y = (i * WOCD)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = "CORNER 4.0 X 5.0 X 16ga"
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", "FRONT", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
      
      P1x = TrailerWidth.Value / 2
      P1y = P1y + 24 - 6
      P2x = TrailerWidth.Value / 2
      P2y = P1y + 6
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Y = (i * WOCD)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = "CORNER 4.0 X 5.0 X 16ga"
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", "FRONT", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
      
            P1x = TrailerWidth.Value / 2
      P1y = TrailerTotalHeight - 6
      P2x = TrailerWidth.Value / 2
      P2y = TrailerTotalHeight
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Y = (i * WOCD)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = "CORNER 4.0 X 5.0 X 16ga"
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", "FRONT", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
      
      ' *********************************************** Front Corner Pieces *****************************************************
      If TrailerWidth.Value = 84 Then
         P1y = -4
      Else
         P1y = -3
      End If
            
      P1x = -TrailerWidth.Value / 2
      P1y = P1y
      P2x = -TrailerWidth.Value / 2
      P2y = P1y + 6
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Y = (i * WOCD)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = "CS_CORNER 4.0 X 5.0 X 16ga"
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", "FRONT", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
      
      P1x = -TrailerWidth.Value / 2
      P1y = P1y + 24 - 6
      P2x = -TrailerWidth.Value / 2
      P2y = P1y + 6
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Y = (i * WOCD)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = "CS_CORNER 4.0 X 5.0 X 16ga"
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", "FRONT", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
      
      P1x = -TrailerWidth.Value / 2
      P1y = TrailerTotalHeight - 6
      P2x = -TrailerWidth.Value / 2
      P2y = TrailerTotalHeight
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Y = (i * WOCD)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = "CS_CORNER 4.0 X 5.0 X 16ga"
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", "FRONT", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
      
      '  ********************************************BACKER MATERIAL FOR RAVEN **********************************
      '  *** Stone Guard Backer ***
      X = (X + 1) * 2
      P1x = -((TrailerWidth.Value / 2) - SW_Offset - (SWFrontVerticalTubeHeight))      ' Allow for .375" Clearance on the Length
      P1y = 24 - MainTubeHeight - 1
      'P2x = ((TrailerWidth.Value / 2) - SW_Offset - (SWFrontVerticalTubeHeight))       ' Added for full width backer (single piece)
      P2x = P1x + WOCD - 0.5 - (SWFrontVerticalTubeHeight)
      P2y = 24 - MainTubeHeight - 1
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = "PLATE 2.0 x 14ga"
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", "FRONT", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
      
      For i = 2 To (X - 1)
 
         P1x = -((TrailerWidth.Value / 2) - SW_Offset) + (i - 1) * WOCD + 0.375    ' Allow for .375" Clearance on the Length
         P1y = 24 - MainTubeHeight - 1
         P2x = P1x + WOCD - 0.75                                              ' TrailerLength.Value - 3.5 - 1.125 - (16)
         P2y = 24 - MainTubeHeight - 1
         ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
         Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
         LineCnt = LineCnt + 1
         'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
         RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
         WeldmentType(LineCnt) = "PLATE 2.0 x 14ga"
         
         Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
         Write #1, "WALL", "FRONT", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
         
      Next i
      
      P1x = ((TrailerWidth.Value / 2) - SW_Offset - (SWFrontVerticalTubeHeight)) - WOCD + 0.0325 + (SWFrontVerticalTubeHeight)      ' Allow for .375" Clearance on the Length
      P1y = 24 - MainTubeHeight - 1
      P2x = ((TrailerWidth.Value / 2) - SW_Offset - (SWFrontVerticalTubeHeight))                         ' TrailerLength.Value - 3.5 - 1.125 - (16)
      P2y = 24 - MainTubeHeight - 1
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = "PLATE 2.0 x 14ga"
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", "FRONT", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
      
      SketchName = Part.SketchManager.ActiveSketch.Name

      '110207 **********************************************************************
      Dim Feature As Object

      Set swApp = Application.SldWorks

      Set Part = swApp.ActiveDoc
      Part.ViewZoomtofit2
      Part.ClearSelection2 True
      boolstatus = Part.Extension.SelectByID2("Point1@Origin", "EXTSKETCHPOINT", 0, 0, 0, True, 6, Nothing, 0)
      longstatus = Part.SketchManager.FullyDefineSketch(1, 1, 1023, 1, 1, Nothing, -1, Nothing, 1, -1)
      Part.ClearSelection2 True
      boolstatus = Part.EditRebuild3
      ' ****************************************************************************

      ' Add Variable for Sketch Plane

      SketchPlane = "FW_SK" ' Specify the SketchName

      ' Call Function to Create Weldments Based on Populated Variables Above
      ReturnVal = BuildWeldment(SketchName, SketchPlane, WeldmentType(), LineCnt, RotationAngle())

   ElseIf NoseStructure.Value = "2ft Wedge" Then               '*****************************2ft Wedge *************************************
      
      If TrailerType.Value = "Motiv Raven Lite Cargo (Steel)" Then 'ListIndex = 0
         If TrailerTotalHeight <= 72 Then
            X = 4
         Else
            X = 5
         End If
         For i = 1 To X
         ReturnVal = RavenWedge(i)
         Next i
      End If
      
      ' *************************** Add Wedge Backers to this Area ******************************
      Dim X_Offset As Double
      
      If TrailerWidth.Value = 60 Then
         X_Offset = 13.3125
      ElseIf TrailerWidth.Value = 72 Then
         X_Offset = 13.3125
      ElseIf TrailerWidth.Value = 84 Then
         X_Offset = 13.3125
      End If
      
      ' Select Working Plane
      boolstatus = Part.Extension.SelectByID2("CSBP", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
      Part.SketchManager.InsertSketch True
      
      LineCnt = 0
      
      P1x = -X_Offset
      P1y = 24
      P2x = -X_Offset
      P2y = TrailerTotalHeight
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = "PLATE 2.0 x 14ga"
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", "FRONT", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
      
      SketchName = Part.SketchManager.ActiveSketch.Name
      
            '110207 **********************************************************************
      'Dim Feature As Object

      Set swApp = Application.SldWorks

      Set Part = swApp.ActiveDoc
      Part.ViewZoomtofit2
      Part.ClearSelection2 True
      boolstatus = Part.Extension.SelectByID2("Point1@Origin", "EXTSKETCHPOINT", 0, 0, 0, True, 6, Nothing, 0)
      longstatus = Part.SketchManager.FullyDefineSketch(1, 1, 1023, 1, 1, Nothing, -1, Nothing, 1, -1)
      Part.ClearSelection2 True
      boolstatus = Part.EditRebuild3
      ' ****************************************************************************

      ' Add Variable for Sketch Plane

      SketchPlane = "CS_WSK" ' Specify the SketchName

      ' Call Function to Create Weldments Based on Populated Variables Above
      ReturnVal = BuildWeldment(SketchName, SketchPlane, WeldmentType(), LineCnt, RotationAngle())
      
      ' **************************************************************************************************
       ' Select Working Plane
      boolstatus = Part.Extension.SelectByID2("RSBP", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
      Part.SketchManager.InsertSketch True
      
      LineCnt = 0
      
      P1x = X_Offset
      P1y = 24
      P2x = X_Offset
      P2y = TrailerTotalHeight
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = "PLATE 2.0 x 14ga"
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", "FRONT", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
      
      SketchName = Part.SketchManager.ActiveSketch.Name
      
            '110207 **********************************************************************
      'Dim Feature As Object

      Set swApp = Application.SldWorks

      Set Part = swApp.ActiveDoc
      Part.ViewZoomtofit2
      Part.ClearSelection2 True
      boolstatus = Part.Extension.SelectByID2("Point1@Origin", "EXTSKETCHPOINT", 0, 0, 0, True, 6, Nothing, 0)
      longstatus = Part.SketchManager.FullyDefineSketch(1, 1, 1023, 1, 1, Nothing, -1, Nothing, 1, -1)
      Part.ClearSelection2 True
      boolstatus = Part.EditRebuild3
      ' ****************************************************************************

      ' Add Variable for Sketch Plane

      SketchPlane = "RS_WSK" ' Specify the SketchName

      ' Call Function to Create Weldments Based on Populated Variables Above
      ReturnVal = BuildWeldment(SketchName, SketchPlane, WeldmentType(), LineCnt, RotationAngle())
   
   End If  ' ******************************* Nose Structure *******************************


      

End Sub


Function BuildWeldment(SketchName As String, SketchPlane As String, WeldmentType() As String, LineCnt As Integer, RotationAngle() As Double)

Dim swApp As Object
Dim Part As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long
Dim i As Integer

Set swApp = Application.SldWorks
Set Part = swApp.ActiveDoc

'110207 Define Sketch Plane To Add Weldments ***********************************************
boolstatus = Part.Extension.SelectByID2(SketchName, "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
boolstatus = Part.Extension.SelectByID2(SketchName, "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
boolstatus = Part.SelectedFeatureProperties(0, 0, 0, 0, 0, 0, 0, 1, 0, SketchPlane)
'*******************************************************************************************

For i = 1 To LineCnt
   ' 110207 Add Weldment to Frame Template ***********************************************************************
      'Set Part = swApp.ActiveDoc
      boolstatus = Part.Extension.SelectByID2("Line" + CStr(i) + "@" + SketchPlane, "EXTSKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0)
      Dim myFeature As Object
      Dim vGroups As Variant
      Dim GroupArray() As Object
      ReDim GroupArray(0 To 0) As Object
      Dim Group1 As Object
      Set Group1 = Part.FeatureManager.CreateStructuralMemberGroup()
      Dim vSegement1 As Variant
      Dim SegementArray1() As Object
      ReDim SegementArray1(0 To 0) As Object
      Part.ClearSelection2 True
      boolstatus = Part.Extension.SelectByID2("Line" + CStr(i) + "@" + SketchPlane, "EXTSKETCHSEGMENT", 0, 0, 0, True, 0, Nothing, 0)
      Dim Segment As Object
      Set Segment = Part.SelectionManager.GetSelectedObject5(1)
      Set SegementArray1(0) = Segment
      vSegement1 = SegementArray1
      Group1.Segments = (vSegement1)
      Group1.ApplyCornerTreatment = True
      Group1.CornerTreatmentType = 1
      Group1.GapWithinGroup = 0
      Group1.GapForOtherGroups = 0
      Group1.Angle = RotationAngle(i)
      Set GroupArray(0) = Group1
      vGroups = GroupArray
      
   ' ****************************************************** TUBE ********************************************************************
   If WeldmentType(i) = "Tube 1.0 x 1.0 x 16ga" Then
      Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "TUBES\1 x 1 x 16ga.sldlfp", 1, True, (vGroups))
      Part.ClearSelection2 True
   End If
   
   If WeldmentType(i) = "Tube 1.0 x 1.5 x 16ga" Then
      Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "TUBES\1 x 1.5 x 16ga.sldlfp", 1, True, (vGroups))
      Part.ClearSelection2 True
   End If
      
   If WeldmentType(i) = "Tube 1.0 x 2.0 x 14ga" Then
      Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "TUBES\1 x 2 x 14ga.sldlfp", 1, True, (vGroups))
      Part.ClearSelection2 True
   End If
   
   If WeldmentType(i) = "Tube 1.0 x 3.0 x 14ga" Then
       Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "TUBES\1 x 3 x 14ga.sldlfp", 1, True, (vGroups))
       Part.ClearSelection2 True
   End If
   
   If WeldmentType(i) = "Tube 2.0 x 3.0 x 11ga" Then
      Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "TUBES\2 x 3 x 11ga.sldlfp", 1, True, (vGroups))
      Part.ClearSelection2 True
   End If
  
   If WeldmentType(i) = "Tube 2.0 x 4.0 x 11ga" Then
       Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "TUBES\2 x 4 x 11ga.sldlfp", 1, True, (vGroups))
       Part.ClearSelection2 True
   End If
    
    If WeldmentType(i) = "Tube 2.0 x 5.0 x 11ga" Then
       Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "TUBES\2 x 5 x 11ga.sldlfp", 1, True, (vGroups))
       Part.ClearSelection2 True
   End If
    
    If WeldmentType(i) = "Tube 2.0 x 8.0 x 11ga" Then
       Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "TUBES\2 x 8 x 11ga.sldlfp", 1, True, (vGroups))
       Part.ClearSelection2 True
   End If
   
   ' ****************************************************** C-CHANNEL ********************************************************************

   If WeldmentType(i) = "C-CHANNEL 1.5 x 2.75 x 13ga" Then
      Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "C-CHANNEL\1.5 x 2.750 x 13 ga C-Channel.sldlfp", 1, True, (vGroups))
      Part.ClearSelection2 True
   End If
   
   If WeldmentType(i) = "C-CHANNEL 1.5 x 3.38 x 13ga" Then
      Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "C-CHANNEL\1.5 x 3.375 x 13 ga C-Channel.sldlfp", 1, True, (vGroups))
      Part.ClearSelection2 True
   End If
      
   If WeldmentType(i) = "C-CHANNEL 1.5 x 3.75 x 13ga" Then
      Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "C-CHANNEL\1.5 x 3.750 x 13 ga C-Channel.sldlfp", 1, True, (vGroups))
      Part.ClearSelection2 True
   End If
      
   If WeldmentType(i) = "C-CHANNEL 1.5 x 4.38 x 13ga" Then
      Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "C-CHANNEL\1.5 x 4.375 x 13 ga C-Channel.sldlfp", 1, True, (vGroups))
      Part.ClearSelection2 True
   End If
      
   If WeldmentType(i) = "C-CHANNEL 1.5 x 4.88 x 13ga" Then
      Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "C-CHANNEL\1.5 x 4.875 x 13 ga C-Channel.sldlfp", 1, True, (vGroups))
      Part.ClearSelection2 True
   End If
      
   If WeldmentType(i) = "C-CHANNEL 1.5 x 5.38 x 13ga" Then
      Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "C-CHANNEL\1.5 x 5.375 x 13 ga C-Channel.sldlfp", 1, True, (vGroups))
      Part.ClearSelection2 True
   End If
         
   If WeldmentType(i) = "C-CHANNEL 1.5 x 8.38 x 13ga" Then
      Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "C-CHANNEL\1.5 x 8.375 x 13 ga C-Channel.sldlfp", 1, True, (vGroups))
      Part.ClearSelection2 True
   End If
   ' ****************************************************** SPECIALTY ********************************************************************
   If WeldmentType(i) = "Z-CHANNEL 1.125 x 1.125" Then
      Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "Z-CHANNEL\Z-CHANNEL 1.125 x 1.125 x 16ga.sldlfp", 1, True, (vGroups))
      Part.ClearSelection2 True
   End If
   
   If WeldmentType(i) = "HEADER 5.00 X 1.25" Then
      Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "Z-CHANNEL\HEADER 5.00 X 1.25 x 14ga.sldlfp", 1, True, (vGroups))
      Part.ClearSelection2 True
   End If
   
   If WeldmentType(i) = "SIDE POST 5.00 X 3.50" Then
      Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "Z-CHANNEL\SIDE POST 5.00 X 3.50 x 14ga.sldlfp", 1, True, (vGroups))
      Part.ClearSelection2 True
   End If
    
   If WeldmentType(i) = "SIDE POST 5.00 X 2.00" Then
      Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "Z-CHANNEL\SIDE POST 5.00 X 2.00 x 14ga.sldlfp", 1, True, (vGroups))
      Part.ClearSelection2 True
   End If
    
   If WeldmentType(i) = "BUMPER 4.00 X 3.38" Then
       Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "Z-CHANNEL\BUMPER 4.00 X 3.38 x 14ga.sldlfp", 1, True, (vGroups))
       Part.ClearSelection2 True
   End If
   
   If WeldmentType(i) = "BUMPER 4.00 X 4.38" Then
       Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "Z-CHANNEL\BUMPER 4.00 X 4.38 x 14ga.sldlfp", 1, True, (vGroups))
       Part.ClearSelection2 True
   End If
   
   ' ****************************************************** ANGLE ********************************************************************
   '1.5 x 1.5 x 0.25.sldlfp
   If WeldmentType(i) = "ANGLE 1.5 X 1.5 x 0.25" Then
       Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "ANGLES\1.5 x 1.5 x 0.25.sldlfp", 1, True, (vGroups))
       Part.ClearSelection2 True
   End If
   ' ****************************************************** PLATE ********************************************************************
  
   If WeldmentType(i) = "PLATE 2.0 x 14ga" Then
       Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "PLATE\2 x 14ga.sldlfp", 1, True, (vGroups))
       Part.ClearSelection2 True
   End If
   
   If WeldmentType(i) = "PLATE 3.0 X 5.0 x 7ga" Then
       Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "PLATE\3 x 7ga.sldlfp", 1, True, (vGroups))
       Part.ClearSelection2 True
   End If
   
   If WeldmentType(i) = "PLATE 4.0 X 7.0 x 7ga" Then
       Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "PLATE\4 x 7ga.sldlfp", 1, True, (vGroups))
       Part.ClearSelection2 True
   End If
   
   If WeldmentType(i) = "PLATE 5.0 X 7.5 x 7ga" Then
       Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "PLATE\5 x 7ga.sldlfp", 1, True, (vGroups))
       Part.ClearSelection2 True
   End If
   
   If WeldmentType(i) = "PLATE 8.0 X 7.5 x 7ga" Then
       Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "PLATE\8 x 7ga.sldlfp", 1, True, (vGroups))
       Part.ClearSelection2 True
   End If
 
   If WeldmentType(i) = "PLATE 8.0 X 14ga" Then
       Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "PLATE\8 x 14ga.SLDLFP", 1, True, (vGroups))
       Part.ClearSelection2 True
   End If
   ' ****************************************************** NEED TO FIX THIS ********************************************************************
   If WeldmentType(i) = "CORNER 4.0 X 5.0 X 16ga" Then
       Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "Z-CHANNEL\CORNER 5.00 X 4.00 x 16ga.sldlfp", 1, True, (vGroups))
       Part.ClearSelection2 True
   End If

   If WeldmentType(i) = "CS_CORNER 4.0 X 5.0 X 16ga" Then
       Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "Z-CHANNEL\CS_CORNER 5.00 X 4.00 x 16ga.sldlfp", 1, True, (vGroups))
       Part.ClearSelection2 True
   End If

   
Next i


End Function

Private Sub CB14_Click()
Dim OCD As Double
Dim TW As Double
Dim Weldment As Double
Dim LineCnt As Integer
Dim X As Integer
Dim i As Integer
' 110207 *************************************************************************************
Dim WOCD As Double
Dim Y As Double
Dim WeldmentType(100) As String
Dim SketchName As String
'Dim SW_Offset As Double

'**********************************************************************************************
Dim swApp As Object

Dim Part As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long
Dim RotationAngle(100) As Double          ' Angle of Rotation of the Weldment Profile

Set swApp = Application.SldWorks
Set Part = swApp.ActiveDoc

'# Setting the Input Parameters
TW = 1
LineCnt = 0

WOCD = 16

   boolstatus = Part.Extension.SelectByID2("Rear Wall@" + CStr(PartNumber.Value) + ".SLDPRT", "CONFIGURATIONS", 0, 0, 0, False, 0, Nothing, 0)
   boolstatus = Part.ShowConfiguration2("Rear Wall")

   '110207 *************************************************************************************
   Part.ClearSelection2 True
   Dim skSegment As Object
   ' *******************************************************************************************

   ' Select Working Plane
   boolstatus = Part.Extension.SelectByID2("RW", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
   Part.SketchManager.InsertSketch True

'# ? Wall Header
P1x = -((TrailerWidth.Value / 2) - 5)
P1y = TrailerTotalHeight - 2.5
P2x = ((TrailerWidth.Value / 2) - 5)
P2y = TrailerTotalHeight - 2.5
ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
LineCnt = LineCnt + 1
'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
WeldmentType(LineCnt) = "HEADER 5.00 X 1.25"

   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", "REAR", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.

'# ? Wall Footer
If FrameMember.Value = "Tube 2.0 x 3.0 x 11ga" Then
   P1x = -(TrailerWidth.Value / 2)
   P1y = -((3.375 / 2) - 0.0625)
   P2x = (TrailerWidth.Value / 2)
   P2y = -((3.375 / 2) - 0.0625)
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   Set skSegment = Part.SketchManager.CreateLine(P2x, P2y, 0#, P1x, P1y, 0#)
   LineCnt = LineCnt + 1
   'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = "BUMPER 4.00 X 3.38"
   
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", "REAR", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
   
ElseIf FrameMember.Value = "Tube 2.0 x 4.0 x 11ga" Then
   P1x = -(TrailerWidth.Value / 2)
   P1y = -((4.375 / 2) - 0.0625)
   P2x = (TrailerWidth.Value / 2)
   P2y = -((4.375 / 2) - 0.0625)
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   Set skSegment = Part.SketchManager.CreateLine(P2x, P2y, 0#, P1x, P1y, 0#)
   LineCnt = LineCnt + 1
   'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = "BUMPER 4.00 X 4.38"
   
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", "REAR", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
   
End If



' CS Vertical Tube

   'Y = (I * WOCD) + FW_Offset
   P1x = -(TrailerWidth.Value / 2)
   P1y = 0.0625
   P2x = -(TrailerWidth.Value / 2)
   P2y = TrailerTotalHeight
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   LineCnt = LineCnt + 1
   'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = "SIDE POST 5.00 X 2.00"

   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", "REAR", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
   
' RS Vertical Tube

'Y = (I * WOCD) + FW_Offset
   P1x = (TrailerWidth.Value / 2)
   P1y = 0.0625
   P2x = (TrailerWidth.Value / 2)
   P2y = TrailerTotalHeight
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   Set skSegment = Part.SketchManager.CreateLine(P2x, P2y, 0#, P1x, P1y, 0#)
   LineCnt = LineCnt + 1
   'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = "SIDE POST 5.00 X 2.00"
   
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", "REAR", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.

'*********************************************************************************

SketchName = Part.SketchManager.ActiveSketch.Name

'110207 **********************************************************************
Dim Feature As Object

Set swApp = Application.SldWorks

Set Part = swApp.ActiveDoc

Part.ViewZoomtofit2
Part.ClearSelection2 True
boolstatus = Part.Extension.SelectByID2("Point1@Origin", "EXTSKETCHPOINT", 0, 0, 0, True, 6, Nothing, 0)
longstatus = Part.SketchManager.FullyDefineSketch(1, 1, 1023, 1, 1, Nothing, -1, Nothing, 1, -1)
Part.ClearSelection2 True
boolstatus = Part.EditRebuild3

' ****************************************************************************
' Add Variable for Sketch Plane
Dim SketchPlane As String

SketchPlane = "RW_SK" ' Specify the SketchName

' Call Function to Create Weldments Based on Populated Variables Above
ReturnVal = BuildWeldment(SketchName, SketchPlane, WeldmentType(), LineCnt, RotationAngle())

End Sub

Private Sub CB2_Click()
   Dim ReturnVal As String
   
   ReturnVal = SideWall("CS")
   DoorOption.Value = 0
   ReturnVal = SideWall("RS")
   
End Sub

Private Sub CB3_Click()
   Dim OCD As Double
   Dim TW As Double
   Dim Weldment As Double
   Dim LineCnt As Integer
   Dim WeldmentType(100) As String
   Dim X As Integer
   Dim i As Integer          ' Counter use in For Statements
   Dim TGP As Double         ' Tongue Gusset Position (Front)
   Dim TGP2 As Double        ' Tongue Gusset Position (Rear - Frame Intersection)

   '110207 *************************************************************************************
   Dim Y As Double
   Dim SketchName As String
   Dim FWLat_Inset As Double
   Dim FWLng_Inset As Double
   ' *******************************************************************************************

   Dim swApp As Object
   Dim Part As Object
   Dim boolstatus As Boolean
   Dim longstatus As Long, longwarnings As Long
   Dim RotationAngle(100) As Double          ' Angle of Rotation of the Weldment Profile
   Dim Dist As Double
   
   Set swApp = Application.SldWorks
   Set Part = swApp.ActiveDoc

   boolstatus = Part.Extension.SelectByID2("Frame@" + CStr(PartNumber.Value) + ".SLDPRT", "CONFIGURATIONS", 0, 0, 0, False, 0, Nothing, 0)
   boolstatus = Part.ShowConfiguration2("Frame")

   '110207 *************************************************************************************
   '# Setting the Input Parameters

   TW = 2  ' Tube Width
   FWLat_Inset = 8 ' Front Wall Laterial Inset for Flat Walls
   FWLng_Inset = 5 ' Front Wall Longitudinal Inset for Flat Walls

   OCD = OCDist.Value
   TGP = TongueLength.Value - ((TGL / 2) / Tan(Theta)) ' Tongue Gusset Position

   ' Select Working Plane
   boolstatus = Part.Extension.SelectByID2("FRAME", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
   Part.SketchManager.InsertSketch True

   '110207 *************************************************************************************
   Part.ClearSelection2 True
   Dim skSegment As Object
   LineCnt = 0
   ' *******************************************************************************************
   ' Tongue Gusset Member Front and Rear
   P1x = (TGL / 2)
   P1y = TGP
   P2x = -(TGL / 2)
   P2y = TGP
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   '# Create Sketch Line
   Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   'Set skSegment = Part.SketchManager.CreateLine(P2x, P2y, 0#, P1x, P1y, 0#)
   LineCnt = LineCnt + 1
   'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   
   If MainTubeHeight = 3 Then   ' ***********************************Calculated Gusset Size Based on Main Rail Size *************************
       WeldmentType(LineCnt) = "PLATE 3.0 X 5.0 x 7ga"
   ElseIf MainTubeHeight = 4 Then
       WeldmentType(LineCnt) = "PLATE 4.0 X 7.0 x 7ga"
   ElseIf MainTubeHeight = 5 Then
       WeldmentType(LineCnt) = "PLATE 5.0 X 7.5 x 7ga"
   ElseIf MainTubeHeight = 8 Then
       WeldmentType(LineCnt) = "PLATE 8.0 X 7.5 x 7ga"
   Else
     MsgBox ("You have a Problem with your tube height, see Frame Sub-Routine")
   End If
   
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "FRAME", "FRAME", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
   
   
   ' Curb Side Gusset Member
   P1x = ((TrailerWidth.Value / 2) - 2)
   P1y = -(((TrailerWidth.Value / 2 - TW - TW / Cos(Theta)) / ((TW / 2) * Tan(Theta))) - TongueLength.Value + (TW / 2) * Sin(Theta) - Sin(Theta) - 2 / Sin(Theta))
   P2x = ((TrailerWidth.Value / 2) - 2) - (2 * Cos(Theta))
   P2y = -(((TrailerWidth.Value / 2 - TW - TW / Cos(Theta)) / ((TW / 2) * Tan(Theta))) - TongueLength.Value + (TW / 2) * Sin(Theta) - Sin(Theta) - 2 / Sin(Theta) + 2 * Sin(Theta))
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   '# Create Sketch Line
   'Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   Set skSegment = Part.SketchManager.CreateLine(P2x, P2y, 0#, P1x, P1y, 0#)
     
   LineCnt = LineCnt + 1
   'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = WeldmentType(LineCnt - 1)
   
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "FRAME", "FRAME", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.

   'Road Side Gusset Member
   P1x = -((TrailerWidth.Value / 2) - 2)
   P1y = -(((TrailerWidth.Value / 2 - TW - TW / Cos(Theta)) / ((TW / 2) * Tan(Theta))) - TongueLength.Value + (TW / 2) * Sin(Theta) - Sin(Theta) - 2 / Sin(Theta))
   P2x = -(((TrailerWidth.Value / 2) - 2) - (2 * Cos(Theta)))
   P2y = -(((TrailerWidth.Value / 2 - TW - TW / Cos(Theta)) / ((TW / 2) * Tan(Theta))) - TongueLength.Value + (TW / 2) * Sin(Theta) - Sin(Theta) - 2 / Sin(Theta) + 2 * Sin(Theta))
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   '# Create Sketch Line
   Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   'Set skSegment = Part.SketchManager.CreateLine(P2x, P2y, 0#, P1x, P1y, 0#)
   LineCnt = LineCnt + 1
   'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = WeldmentType(LineCnt - 1)
   
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "FRAME", "FRAME", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.

   If NoseStructure.Value = "2ft Wedge" Then ' **************************** 2 ft Wedge ************************************

      FWLat_Inset = 0
      FWLng_Inset = 0

      P1x = (Tan(Theta) * (TongueLength.Value - 13.75))
      P1y = 13.75
      P2x = -(Tan(Theta) * (TongueLength.Value - 13.75))
      P2y = 13.75
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      'Set skSegment = Part.SketchManager.CreateLine(P2x, P2y, 0#, P1x, P1y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = InsideChannel
   
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "FRAME", "FRAME", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
   
   ElseIf NoseStructure.Value = "4ft Wedge" Then ' **************************** 4 ft Wedge ************************************
      'MsgBox ("Not Yet")
      End
   ElseIf NoseStructure.Value = "6ft Wedge" Then ' **************************** 6 ft Wedge ************************************
      'MsgBox ("Not Yet")
      End
   Else
      '# 5 Front Stub Cross Member (Only on Flat Nose Trailers)

      P1x = TrailerWidth.Value / 2
      P1y = -5
      P2x = (TongueLength.Value + 5) * Tan(Theta) + 2 / Cos(Theta)
      P2y = -5
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = OverChannel

      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "FRAME", "FRAME", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.

      P1x = -TrailerWidth.Value / 2
      P1y = -5
      P2x = -((TongueLength.Value + 5) * Tan(Theta) + 2 / Cos(Theta))
      P2y = -5
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      ' Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      Set skSegment = Part.SketchManager.CreateLine(P2x, P2y, 0#, P1x, P1y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = OverChannel
   
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "FRAME", "FRAME", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
   
      '# 1 Front Stub Laterial Members

      P1x = ((TrailerWidth.Value - FWLat_Inset) / 2) - 1
      P1y = -0.0897
      P2x = ((TrailerWidth.Value - FWLat_Inset) / 2) - 1
      P2y = -5
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = FrameMember.Value
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "FRAME", "FRAME", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
      
      P1x = -(((TrailerWidth.Value - FWLat_Inset) / 2) - 1)
      P1y = -0.0897
      P2x = -(((TrailerWidth.Value - FWLat_Inset) / 2) - 1)
      P2y = -5
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = FrameMember.Value
   
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "FRAME", "FRAME", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
   
   End If

   '# 6 Front Cross Member
   P1x = -(TrailerWidth.Value - FWLat_Inset) / 2
   P1y = -0
   P2x = (TrailerWidth.Value - FWLat_Inset) / 2
   P2y = -0
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)

   'Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   Set skSegment = Part.SketchManager.CreateLine(P2x, P2y, 0#, P1x, P1y, 0#)
   
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "FRAME, FRAME,"; OverChannel, Dist, 0, 0  ' Write comma-delimited data.
   
   LineCnt = LineCnt + 1
   'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = OverChannel

   '# 7 Cross Members
   '# Note: Repeat this Operation based on Length and Frame Center to Center Distance

   X = (TrailerLength.Value / OCD) - 1

   For i = 1 To X
      Y = i * OCD
      If TrailerWidth.Value > 80 And Y > (AxleCL - (SWAxleClearance / 2) - SWRailOverlap) And Y < (AxleCL + (SWAxleClearance / 2) + SWRailOverlap) Then
         P1x = -((TrailerWidth.Value - 4) / 2 - TW)
         P1y = -Y
         P2x = ((TrailerWidth.Value - 4) / 2 - TW)
         P2y = -Y
         ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      ElseIf Y < (((TrailerWidth.Value / 2 - TW - TW / Cos(Theta)) / ((TW / 2) * Tan(Theta))) - TongueLength.Value + (TW / 2) * Sin(Theta)) Then
         P1x = -(Tan(Theta) * (TongueLength.Value + Y))
         P1y = -Y
         P2x = (Tan(Theta) * (TongueLength.Value + Y))
         P2y = -Y
         ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
         
      Else
         P1x = -(TrailerWidth.Value - 4) / 2
         P1y = -Y
         P2x = (TrailerWidth.Value - 4) / 2
         P2y = -Y
         ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      End If
      
      'Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      Set skSegment = Part.SketchManager.CreateLine(P2x, P2y, 0#, P1x, P1y, 0#)
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "FRAME, FRAME,"; InsideChannel, Dist, 0, 0  ' Write comma-delimited data.
      
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = InsideChannel
   
      If Y < ((((TrailerWidth.Value / 2 - TW - TW / Cos(Theta)) / ((TW / 2) * Tan(Theta))) - TongueLength.Value + (TW / 2) * Sin(Theta)) - 12) Then
           
            ' RS Short Laterial Piece
            P1x = -(TrailerWidth.Value - 4) / 2
            P1y = -(Y + 1.5)
            P2x = -((Tan(Theta) * (TongueLength.Value + Y)) + (TW / Cos(Theta) + Tan(Theta) * 1.5))
            P2y = -(Y + 1.5)
            ReturnVal = Convert2Meters(P2x, P2y, P1x, P1y)
            
            'Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
            Set skSegment = Part.SketchManager.CreateLine(P2x, P2y, 0#, P1x, P1y, 0#)
            
            Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
            Write #1, "FRAME, FRAME,"; InsideChannel, Dist, 0, 0  ' Write comma-delimited data.
            
            LineCnt = LineCnt + 1
            'RotationAngle(LineCnt) = 1.570796326795          '  90 Degree Rotation Angle
            RotationAngle(LineCnt) = 3.141592654             ' 180 Degree Rotation Angle
            'RotationAngle(LineCnt) = 0                        '   0 Degree Rotation Angle
            WeldmentType(LineCnt) = InsideChannel
            
            ' CS Short Laterial Piece
            P1x = ((Tan(Theta) * (TongueLength.Value + Y)) + (TW / Cos(Theta) + Tan(Theta) * 1.5))
            P1y = -(Y + 1.5)
            P2x = (TrailerWidth.Value - 4) / 2
            P2y = -(Y + 1.5)
            ReturnVal = Convert2Meters(P2x, P2y, P1x, P1y)
                        
            'Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
            Set skSegment = Part.SketchManager.CreateLine(P2x, P2y, 0#, P1x, P1y, 0#)
            
            Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
            Write #1, "FRAME, FRAME,"; InsideChannel, Dist, 0, 0  ' Write comma-delimited data.
            
            LineCnt = LineCnt + 1
            'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
            RotationAngle(LineCnt) = 3.141592654             ' 180 Degree Rotation Angle
            'RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
            WeldmentType(LineCnt) = InsideChannel
         Else
         End If
   Next i

   '# 3 Main Rails

   If TrailerWidth.Value > 80 Then

      P1x = -((TrailerWidth.Value / 2) - 1) ' Front Piece RS
      P1y = -(FWLng_Inset + 0.0897)
      P2x = -((TrailerWidth.Value / 2) - 1)
      P2y = -(AxleCL - (SWAxleClearance / 2))
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "FRAME, FRAME,"; FrameMember.Value, Dist, 0, 0  ' Write comma-delimited data.
      
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = FrameMember.Value
   
      P1x = -((TrailerWidth.Value / 2) - 1 - TW) ' Second Piece (inset) RS
      P1y = -(AxleCL - (SWAxleClearance / 2) - SWRailOverlap)
      P2x = -((TrailerWidth.Value / 2) - 1 - TW)
      P2y = -(AxleCL + (SWAxleClearance / 2) + SWRailOverlap)
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "FRAME, FRAME,"; FrameMember.Value, Dist, 0, 0  ' Write comma-delimited data.
      
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = FrameMember.Value
   
      P1x = -((TrailerWidth.Value / 2) - 1) ' Third Piece RS
      P1y = -(AxleCL + (SWAxleClearance / 2))
      P2x = -((TrailerWidth.Value / 2) - 1)
      P2y = -(TrailerLength.Value)
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "FRAME, FRAME,"; FrameMember.Value, Dist, 0, 0  ' Write comma-delimited data.
      
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = FrameMember.Value

      P1x = ((TrailerWidth.Value / 2) - 1) ' First Piece CS
      P1y = -(FWLng_Inset + 0.0897)
      P2x = ((TrailerWidth.Value / 2) - 1)
      P2y = -(AxleCL - (SWAxleClearance / 2))
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "FRAME, FRAME,"; FrameMember.Value, Dist, 0, 0  ' Write comma-delimited data.
      
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = FrameMember.Value
   
      P1x = (TrailerWidth.Value / 2) - 1 - TW ' Second Piece CS
      P1y = -(AxleCL - (SWAxleClearance / 2) - SWRailOverlap)
      P2x = (TrailerWidth.Value / 2) - 1 - TW
      P2y = -(AxleCL + (SWAxleClearance / 2) + SWRailOverlap)
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "FRAME, FRAME,"; FrameMember.Value, Dist, 0, 0  ' Write comma-delimited data.
      
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = FrameMember.Value
   
      P1x = ((TrailerWidth.Value / 2) - 1)  ' Third Piece CS
      P1y = -(AxleCL + (SWAxleClearance / 2))
      P2x = ((TrailerWidth.Value / 2) - 1)
      P2y = -(TrailerLength.Value)
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "FRAME, FRAME,"; FrameMember.Value, Dist, 0, 0  ' Write comma-delimited data.
      
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = FrameMember.Value
   
   Else ' Trailer Width Less Then 80 Inches
      P1x = ((TrailerWidth.Value / 2) - 1)
      P1y = -FWLng_Inset - 0.0897
      P2x = ((TrailerWidth.Value / 2) - 1)
      P2y = -TrailerLength.Value
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "FRAME, FRAME,"; FrameMember.Value, Dist, 0, 0  ' Write comma-delimited data.
      
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = FrameMember.Value

      P1x = -((TrailerWidth.Value / 2) - 1)
      P1y = -FWLng_Inset - 0.0897
      P2x = -((TrailerWidth.Value / 2) - 1)
      P2y = -TrailerLength.Value
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "FRAME, FRAME,"; FrameMember.Value, Dist, 0, 0  ' Write comma-delimited data.
      
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = FrameMember.Value
   End If

   '# 2 Tongue (A-Frame)

   P1x = -((TW / 2) * Cos(Theta))
   P1y = (TongueLength.Value + (TW / 2) * Sin(Theta))
   P2x = -(TrailerWidth.Value / 2 - TW - (TW / 2) * Cos(Theta))
   P2y = -(((TrailerWidth.Value / 2 - TW - TW / Cos(Theta)) / ((TW / 2) * Tan(Theta))) - TongueLength.Value + (TW / 2) * Sin(Theta))
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "FRAME, FRAME,"; FrameMember.Value, Dist, 0, 0  ' Write comma-delimited data.
   
   LineCnt = LineCnt + 1
   'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = FrameMember.Value   ' Need to Fix This

   P1x = (TW / 2) * Cos(Theta)
   P1y = (TongueLength.Value + (TW / 2) * Sin(Theta))
   P2x = (TrailerWidth.Value / 2 - TW - (TW / 2) * Cos(Theta))
   P2y = -(((TrailerWidth.Value / 2 - TW - TW / Cos(Theta)) / ((TW / 2) * Tan(Theta))) - TongueLength.Value + (TW / 2) * Sin(Theta))
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "FRAME, FRAME,"; FrameMember.Value, Dist, 0, 0  ' Write comma-delimited data.
   
   LineCnt = LineCnt + 1
   'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = FrameMember.Value   ' Need to Fix This

   '110207 **********************************************************************
   Dim Feature As Object

   SketchName = Part.SketchManager.ActiveSketch.Name

   Set swApp = Application.SldWorks
   Set Part = swApp.ActiveDoc
   Part.ViewZoomtofit2
   Part.ClearSelection2 True
   boolstatus = Part.Extension.SelectByID2("Point1@Origin", "EXTSKETCHPOINT", 0, 0, 0, True, 6, Nothing, 0)
   longstatus = Part.SketchManager.FullyDefineSketch(1, 1, 1023, 1, 1, Nothing, -1, Nothing, 1, -1)
   Part.ClearSelection2 True
   boolstatus = Part.EditRebuild3
   ' ****************************************************************************

   ' Add Variable for Sketch Plane
   Dim SketchPlane As String

   SketchPlane = "FRAME_SK" ' Specify the SketchName

   ' Call Function to Create Weldments Based on Populated Variables Above
   ReturnVal = BuildWeldment(SketchName, SketchPlane, WeldmentType(), LineCnt, RotationAngle())

End Sub

Private Sub CB1_Click()

   Dim TW As Double
   Dim Weldment As Double
   Dim LineCnt As Integer
   Dim X As Integer
   Dim i As Integer
   ' 110207 *************************************************************************************
   Dim WOCD As Double
   Dim Y As Double
   Dim WeldmentType(100) As String
   Dim SketchName As String

    ' Add Variable for Sketch Plane
   Dim SketchPlane As String
   '**********************************************************************************************

   Dim swApp As Object
   Dim Part As Object
   Dim boolstatus As Boolean
   Dim longstatus As Long, longwarnings As Long
   Dim RotationAngle(100) As Double          ' Angle of Rotation of the Weldment Profile

   Set swApp = Application.SldWorks
   Set Part = swApp.ActiveDoc

   '# Setting the Input Parameters
   TW = 1
   LineCnt = 0
   WOCD = 24
   RampTubeWidth = 1

   boolstatus = Part.Extension.SelectByID2("Ramp@" + CStr(PartNumber.Value) + ".SLDPRT", "CONFIGURATIONS", 0, 0, 0, False, 0, Nothing, 0)
   boolstatus = Part.ShowConfiguration2("Ramp")
   
   '110207 *************************************************************************************
   Part.ClearSelection2 True
   Dim skSegment As Object
   ' *******************************************************************************************

   ' Select Working Plane
   boolstatus = Part.Extension.SelectByID2("RW", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
   Part.SketchManager.InsertSketch True
      
   '# Wall Header **********************************************************
   P1x = -((TrailerWidth.Value / 2) - 5.5)
   P1y = TrailerTotalHeight - 5 - 1
   P2x = ((TrailerWidth.Value / 2) - 5.5)
   P2y = TrailerTotalHeight - 5 - 1
   ' Convert to Meters
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   LineCnt = LineCnt + 1
   RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   'RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"
   
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", "RAMP", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.

   '# Wall Footer #1 ***********************************************************
   P1x = -((TrailerWidth.Value / 2) - 5.5)
   P1y = 1.0625
   P2x = ((TrailerWidth.Value / 2) - 5.5)
   P2y = 1.0625
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   LineCnt = LineCnt + 1
   RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   'RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"
   
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", "RAMP", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
   
   '# Wall Footer #2 ***********************************************************
   P1x = -((TrailerWidth.Value / 2) - 5.5)
   P1y = 1 + 1.0625
   P2x = ((TrailerWidth.Value / 2) - 5.5)
   P2y = 1 + 1.0625
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   LineCnt = LineCnt + 1
   RotationAngle(LineCnt) = 1.570796326795            ' 90 Degree Rotation Angle
   'RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"
   
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", "RAMP", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
   
   '# Wall Footer #3 Angle ***********************************************************
   P1x = -((TrailerWidth.Value / 2) - 5.5)
   P1y = -0.25 + 0.0625
   P2x = ((TrailerWidth.Value / 2) - 5.5)
   P2y = -0.25 + 0.0625
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   LineCnt = LineCnt + 1
   RotationAngle(LineCnt) = 1.570796326795           ' 90 Degree Rotation Angle
   'RotationAngle(LineCnt) = 4.71238898               ' 270 Degree Rotation Angle
   'RotationAngle(LineCnt) = 3.14159265                ' 270 Degree Rotation Angle
   'RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = "ANGLE 1.5 X 1.5 x 0.25"
   
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", "RAMP", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.

   ' CS Vertical Tube ********************************************************
   
   P1x = -((TrailerWidth.Value / 2) - 6)
   P1y = 2.5625
   P2x = -((TrailerWidth.Value / 2) - 6)
   P2y = TrailerTotalHeight - 5 - 1 - 0.5
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   Y = (i * WOCD)
   Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   LineCnt = LineCnt + 1
   RotationAngle(LineCnt) = 1.570796326795            ' 90 Degree Rotation Angle
   'RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"
   
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", "RAMP", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
   
   ' CS Vertical Tube 12" to Side ********************************************************
   
   P1x = -((TrailerWidth.Value / 2) - 6 - 12 - 1)
   P1y = 2.5625
   P2x = -((TrailerWidth.Value / 2) - 6 - 12 - 1)
   P2y = TrailerTotalHeight - 5 - 1 - 0.5
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   Y = (i * WOCD)
   Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   LineCnt = LineCnt + 1
   RotationAngle(LineCnt) = 1.570796326795            ' 90 Degree Rotation Angle
   'RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"
   
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", "RAMP", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.

   ' RS Vertical Tube *********************************************************
   P1x = ((TrailerWidth.Value / 2) - 6)
   P1y = 2.5625
   P2x = ((TrailerWidth.Value / 2) - 6)
   P2y = TrailerTotalHeight - 5 - 1 - 0.5
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   Y = (i * WOCD)
   Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   LineCnt = LineCnt + 1
   RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   'RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"
   
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", "RAMP", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
   
      ' RS Vertical Tube : Grab Handle Support  *********************************************************
   P1x = ((TrailerWidth.Value / 2) - 6 - RampTubeWidth)
   P1y = TrailerTotalHeight / 2 - 2.75 + 4
   P2x = ((TrailerWidth.Value / 2) - 6 - RampTubeWidth)
   P2y = P1y + 16
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   'Y = (i * WOCD)
   Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   LineCnt = LineCnt + 1
   RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   'RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"
   
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", "RAMP", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
   
   ' RS Vertical Tube 12" to Side *********************************************************
   P1x = ((TrailerWidth.Value / 2) - 6 - 12 - 1)
   P1y = 2.5625
   P2x = ((TrailerWidth.Value / 2) - 6 - 12 - 1)
   P2y = TrailerTotalHeight - 5 - 1 - 0.5
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   Y = (i * WOCD)
   Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   LineCnt = LineCnt + 1
   RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   'RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"
   
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", "RAMP", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
   
   ' Raven Front Only
   ' Center Tube **********************************************************

   P1x = 0
   P1y = 2.5625
   P2x = 0
   P2y = TrailerTotalHeight - 5 - 1 - 0.5
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   Y = (i * WOCD)
   Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   LineCnt = LineCnt + 1
   RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   'RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"
   '# Note: Repeat this Operation based on Length and Frame Center to Center Distance
   
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", "RAMP", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
      
   '  ********************************************BACKER MATERIAL FOR RAVEN **********************************
   '  *** Stone Guard Backer ***
      
   P1x = ((TrailerWidth.Value / 2) - 6.5)      ' Allow for .375" Clearance on the Length
   P1y = 10
   P2x = ((TrailerWidth.Value / 2) - 6.5 - 12)      ' TrailerLength.Value - 3.5 - 1.125 - (16)
   P2y = 10
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   LineCnt = LineCnt + 1
   'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = "PLATE 2.0 x 14ga"
   
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", "RAMP", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
      
   P1x = -((TrailerWidth.Value / 2) - 6.5)      ' Allow for .375" Clearance on the Length
   P1y = TrailerTotalHeight / 2 - 2.75
   P2x = -((TrailerWidth.Value / 2) - 6.5 - 12)      ' TrailerLength.Value - 3.5 - 1.125 - (16)
   P2y = TrailerTotalHeight / 2 - 2.75
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   LineCnt = LineCnt + 1
   RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   'RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = "PLATE 8.0 X 14ga"
   
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", "RAMP", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
   
   P1x = ((TrailerWidth.Value / 2) - 6.5)      ' Allow for .375" Clearance on the Length
   P1y = TrailerTotalHeight / 2 - 2.75
   P2x = ((TrailerWidth.Value / 2) - 6.5 - 12)      ' TrailerLength.Value - 3.5 - 1.125 - (16)
   P2y = TrailerTotalHeight / 2 - 2.75
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   LineCnt = LineCnt + 1
   RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   'RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = "PLATE 8.0 X 14ga"
   
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", "RAMP", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
      
   SketchName = Part.SketchManager.ActiveSketch.Name

   '110207 **********************************************************************
   Dim Feature As Object
   Set swApp = Application.SldWorks

   Set Part = swApp.ActiveDoc
   Part.ViewZoomtofit2
   Part.ClearSelection2 True
   boolstatus = Part.Extension.SelectByID2("Point1@Origin", "EXTSKETCHPOINT", 0, 0, 0, True, 6, Nothing, 0)
   longstatus = Part.SketchManager.FullyDefineSketch(1, 1, 1023, 1, 1, Nothing, -1, Nothing, 1, -1)
   Part.ClearSelection2 True
   boolstatus = Part.EditRebuild3
   ' ****************************************************************************

   ' Add Variable for Sketch Plane

   SketchPlane = "RAMP_SK" ' Specify the SketchName

   ' Call Function to Create Weldments Based on Populated Variables Above
   ReturnVal = BuildWeldment(SketchName, SketchPlane, WeldmentType(), LineCnt, RotationAngle())

End Sub

Private Sub CB5_Click()
Dim swApp As Object

Dim Part As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long


PartDirectoryPath = "C:\Users\markc\Documents\PROJECTS\RAVENS\"
Dim CurrentPart As String
Close #1
'Open "C:\Users\markc\Documents\PROJECTS\RAVENS1\" + CStr(PartNumber.Value) + ".txt" For Output As 1
Open "C:\Users\markc\Documents\Engineering\" + CStr(PartNumber.Value) + ".txt" For Output As 1

Write #1, CStr(PartNumber.Value)    ' Write comma-delimited data.
Write #1,                           ' Write blank line.

Set swApp = Application.SldWorks
Set Part = swApp.NewDocument("C:\ATC ENGINEERING\FORMATS\Start Part.prtdot", 0, 0, 0)
Set Part = swApp.ActiveDoc
longstatus = Part.SaveAs3(PartDirectoryPath + CStr(PartNumber.Value) + ".SLDPRT", 0, 0)

swApp.SetUserPreferenceToggle swUserPreferenceToggle_e.swSketchInference, False

Theta = 0.436332 ' *** Set Angle of the Tongue to 25 Degree *** Value is in Radians

ReturnVal = TubeSize()
ReturnVal = SideWallTubeSize()
ReturnVal = CreatePlanes()
'ReturnVal = CreateConfigurations()
ReturnVal = CreateConfigurations2()
'ReturnVal = CreateDerivedConfig()
'ReturnVal = yetagain()
'ReturnVal = ConfigProp()
ReturnVal = TongueDetails()
End Sub

Private Sub CB6_Click()

Dim swApp As Object
Dim Part As Object
Dim SelMgr As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long
Dim Feature As Object

Set swApp = Application.SldWorks

Set Part = swApp.ActiveDoc
Set SelMgr = Part.SelectionManager
Part.FeatureManager.EditRollback swMoveRollbackBarToBeforeFeature, ""

Set swApp = Application.SldWorks

Set Part = swApp.ActiveDoc
Set SelMgr = Part.SelectionManager
boolstatus = Part.EditRebuild3
Part.ClearSelection2 True


End Sub

Private Sub CB7_Click()

End Sub

Private Sub CB8_Click()
Dim swApp As Object
Dim Part As Object
Dim SelMgr As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long
Dim Feature As Object


Dim swApp As SldWorks.SldWorks

Dim Part As SldWorks.ModelDoc2


'Set SwApp = Application.SldWorks

'Set Part = SwApp.ActiveDoc
Set SelMgr = Part.SelectionManager
Part.ViewZoomtofit2
Part.ClearSelection2 True
boolstatus = Part.Extension.SelectByID2("Point1@Origin", "EXTSKETCHPOINT", 0, 0, 0, True, 6, Nothing, 0)
longstatus = Part.SketchManager.FullyDefineSketch(1, 1, 1023, 1, 1, Nothing, -1, Nothing, 1, -1)
Part.ClearSelection2 True
boolstatus = Part.EditRebuild3
End Sub

Private Sub CB9_Click()

   Dim OCD As Double
   Dim TW As Double
   Dim Weldment As Double
   Dim LineCnt As Integer
   Dim X As Integer
   Dim i As Integer
   ' 110207 *************************************************************************************
   Dim WOCD As Double
   Dim Y As Double
   Dim WeldmentType(100) As String
   Dim SketchName As String
   '**********************************************************************************************
   Dim DoorTubeHeight As Double
   Dim LAST_VERTICAL_POST As Double
   Dim swApp As Object
   
   Dim Part As Object
   Dim boolstatus As Boolean
   Dim longstatus As Long, longwarnings As Long
   Dim RotationAngle(100) As Double              ' Angle of Rotation of the Weldment Profile
   Dim VentFlag As Integer                       ' Flag for Location of Vent

   Set swApp = Application.SldWorks
   Set Part = swApp.ActiveDoc

   '# Setting the Input Parameters

   TW = 1         ' Tube Width ????  ************************** Fix This **********************************
   LineCnt = 0
   VentFlag = 0
   WOCD = 24
   DoorTubeHeight = 2
   RoofTubeWidth = 1

   boolstatus = Part.Extension.SelectByID2("Roof@" + CStr(PartNumber.Value) + ".SLDPRT", "CONFIGURATIONS", 0, 0, 0, False, 0, Nothing, 0)
   boolstatus = Part.ShowConfiguration2("ROOF")

   '110207 *************************************************************************************
   Part.ClearSelection2 True
   Dim skSegment As Object
   ' *******************************************************************************************

   ' Select Working Plane
   boolstatus = Part.Extension.SelectByID2("ROOF", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
   Part.SketchManager.InsertSketch True

   If TrailerType.Value = "Motiv Raven Lite Cargo (Steel)" Then ' ******************************** Motiv Raven Lite Cargo (Steel) ************************************

   If NoseStructure.Value = "Flat Nose" Then ' ******************************* Nose Structure: Flat *******************************

      ' ****** Roof to Side Wall Support Pieces ******
      
      P1y = -((RoofTubeWidth))
      P1x = -(TrailerWidth.Value / 2 - 7.5)
      P2y = -(FW_Offset)
      P2x = -(TrailerWidth.Value / 2 - 7.5)
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"
   
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", "ROOF", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
   
      P1y = -((RoofTubeWidth))
      P1x = (TrailerWidth.Value / 2 - 7.5)
      P2y = -(FW_Offset)
      P2x = (TrailerWidth.Value / 2 - 7.5)
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)

      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"
   
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", "ROOF", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
      End If
   

   '# ****** Vertical Members ******
   '# Note: Repeat this Operation based on Length and Frame Center to Center Distance


   X = (TrailerLength.Value / WOCD)
   
   P1y = -(Y + FW_Offset + (RoofTubeWidth / 2))
   P1x = -(TrailerWidth.Value / 2 - RoofTubeWidth)
   P2y = -(Y + FW_Offset + (RoofTubeWidth / 2))
   P2x = (TrailerWidth.Value / 2 - RoofTubeWidth)
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   Y = (WOCD)
   Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   
   LineCnt = LineCnt + 1
   RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"
   
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", "ROOF", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
   
   For i = 2 To X
   
         ' ****** Add Framing for Vent ******
      If TrailerWidth.Value = 60 And i = 2 Then
         VentFlag = 1
      ElseIf TrailerWidth.Value > 60 And i = 3 Then
         VentFlag = 1
      Else
         VentFlag = 0
      End If
      
      If VentFlag = 1 Then
        
         VentFlag = 0
         
         P1y = -(Y + 0.035 + FW_Offset + RoofTubeWidth / 2)
         P1x = -(14.25 / 2 + RoofTubeWidth / 2)
         P2y = -(Y + 0.035 + FW_Offset + (WOCD) - RoofTubeWidth / 2)
         P2x = -(14.25 / 2 + RoofTubeWidth / 2)
         ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
         'LAST_VERTICAL_POST = Y + 0.035 + FW_Offset
         'Y = (i * WOCD)
         Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
         LineCnt = LineCnt + 1
         RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
         WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"
      
         Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
         Write #1, "WALL", "ROOF", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
      
         P1y = -(Y + 0.035 + FW_Offset + RoofTubeWidth / 2)
         P1x = (14.25 / 2 + RoofTubeWidth / 2)
         P2y = -(Y + 0.035 + FW_Offset + (WOCD) - RoofTubeWidth / 2)
         P2x = (14.25 / 2 + RoofTubeWidth / 2)
         ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
         'LAST_VERTICAL_POST = Y + 0.035 + FW_Offset
         'Y = (i * WOCD)
         Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
         LineCnt = LineCnt + 1
         RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
         WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"
      
         Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
         Write #1, "WALL", "ROOF", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
      
         '********************************************* Cross Piece
         P1y = -(Y + 0.035 + FW_Offset + 14.25 + RoofTubeWidth)
         P1x = -(14.25 / 2)
         P2y = -(Y + 0.035 + FW_Offset + 14.25 + RoofTubeWidth)
         P2x = (14.25 / 2)
         ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
         'LAST_VERTICAL_POST = Y + 0.035 + FW_Offset
         'Y = (i * WOCD)
         Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
         LineCnt = LineCnt + 1
         RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
         WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"
      
         Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
         Write #1, "WALL", "ROOF", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
         
      End If
      
      P1y = -(Y + 0.035 + FW_Offset)
      P1x = -(TrailerWidth.Value / 2 - RoofTubeWidth)
      P2y = -(Y + 0.035 + FW_Offset)
      P2x = (TrailerWidth.Value / 2 - RoofTubeWidth)
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      LAST_VERTICAL_POST = Y + 0.035 + FW_Offset
      Y = (i * WOCD)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", "ROOF", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
      
   Next i

   P1y = -(TrailerLength.Value - 3.5 - RoofTubeWidth / 2)
   P1x = -(TrailerWidth.Value / 2 - RoofTubeWidth)
   P2y = -(TrailerLength.Value - 3.5 - RoofTubeWidth / 2)
   P2x = (TrailerWidth.Value / 2 - RoofTubeWidth)
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   Y = (i * WOCD)
   Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)

   LineCnt = LineCnt + 1
   RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"
         
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", "ROOF", WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
   
   ElseIf TrailerType.Value = "Motiv MSX Series" Then  ' ******************************** Motiv MSX Series ************************************
      MsgBox ("Not Yet")
      End
   ElseIf TrailerType.Value = "Motiv RSX Gooseneck Car Hauler" Then
      MsgBox ("Not Yet")
      End
   ElseIf TrailerType.Value = "Motiv RSX Car Hauler" Then
      MsgBox ("Not Yet")
      End
   ElseIf TrailerType.Value = "Motiv SSX Series Stacker" Then
      MsgBox ("Not Yet")
      End
   ElseIf TrailerType.Value = "Motiv Steel Snow Trailer" Then
      MsgBox ("Not Yet")
      End
   End If

   SketchName = Part.SketchManager.ActiveSketch.Name

   '110207 **********************************************************************
   Dim Feature As Object

   Set swApp = Application.SldWorks

   Set Part = swApp.ActiveDoc
   Part.ViewZoomtofit2
   Part.ClearSelection2 True
   boolstatus = Part.Extension.SelectByID2("Point1@Origin", "EXTSKETCHPOINT", 0, 0, 0, True, 6, Nothing, 0)
   longstatus = Part.SketchManager.FullyDefineSketch(1, 1, 1023, 1, 1, Nothing, -1, Nothing, 1, -1)
   Part.ClearSelection2 True
   boolstatus = Part.EditRebuild3

   ' ****************************************************************************
   ' Add Variable for Sketch Plane
   Dim SketchPlane As String

   SketchPlane = "ROOF_SK" ' Specify the SketchName

   ' Call Function to Create Weldments Based on Populated Variables Above
   ReturnVal = BuildWeldment(SketchName, SketchPlane, WeldmentType(), LineCnt, RotationAngle())

End Sub

Private Sub CommandButton2_Click()

UserForm2.Show

End Sub


Private Sub UserForm_Initialize()

ReturnVal = SetMenu()
 
End Sub

Function SetMenu()
WeldmentPath = "C:\ATC ENGINEERING\EXTRUSIONS\STEEL\"

' Definitions:
' ***************************
'    Frame Member:
'       Tube 2.0 x 3.0 x 11ga
'       Tube 2.0 x 4.0 x 11ga
'       Tube 2.0 x 5.0 x 11ga
'       Tube 2.0 x 8.0 x 11ga
'       C-CHANNEL 1.5 x 2.75 x 13ga
'       C-CHANNEL 1.5 x 3.38 x 13ga
'       C-CHANNEL 1.5 x 3.75 x 13ga
'       C-CHANNEL 1.5 x 4.38 x 13ga
'       C-CHANNEL 1.5 x 4.75 x 13ga
'       C-CHANNEL 1.5 x 5.38 x 13ga
'       C-CHANNEL 1.5 x 8.38 x 13ga
'       PLATE 2.0 x 14ga
'       PLATE 3.0 X 5.0 x 7ga
'       PLATE 4.0 X 7.0 x 7ga
'       PLATE 5.0 X 7.5 x 7ga
'       PLATE 8.0 X 7.5 x 7ga
'       PLATE 8.0 X 14ga
'       ANGLE 1.5 X 1.5 x 0.25
' ***************************
'   Side Wall Member
'      Z-CHANNEL 1.125 x 1.125
'      Tube 1.0 x 1.0 x 16ga
'      Tube 1.0 x 1.5 x 16ga
'      Tube 1.0 x 2.0 x 14ga
'      Tube 1.0 x 3.0 x 14ga
'      Tube 2.0 x 2.0 x 11ga
' ***************************
'   Front Wall Structure:
'      Flat Nose
'      2ft Wedge
'      4ft Wedge
'      6ft Wedge
'      CORNER 4.0 X 5.0 X 16ga
'      CS_CORNER 4.0 X 5.0 X 16ga
' ***************************
'    Rear Wall Structure:
'       HEADER 5.00 X 1.25
'       SIDE POST 5.00 X 3.50
'       SIDE POST 5.00 X 2.00
'       BUMPER 4.00 X 3.38
'       BUMPER 4.00 X 4.38
' ***************************
'   Trailer Type:
'      Motiv Raven Lite Cargo (Steel)
'      Motiv MSX Series
'      Motiv RSX Gooseneck Car Hauler
'      Motiv RSX Car Hauler
'      Motiv SSX Series Stacker
'      Motiv Steel Snow Trailer
'
'      Motiv Raven Lite Cargo (Aluminum)
'      ATC Quest Car Hauler
'      ATC Quest Gooseneck Car Hauler
'      ATC SS Series Trailer
'      ATC Signature Series Stacker
'      ATC Northstar Snowmobile Trailer
'      ATC Concours Series Open Car Hauler

'Add list entries to combo box. The value of each
    'entry matches the corresponding ListIndex value
    'in the combo box.
    TrailerWidth.AddItem "60"                            'ListIndex = 0
    TrailerWidth.AddItem "72"                            'ListIndex = 1
    TrailerWidth.AddItem "84"                            'ListIndex = 2
    TrailerWidth.AddItem "96"                            'ListIndex = 3
    TrailerWidth.AddItem "100"                           'ListIndex = 4

    TrailerLength.AddItem "98"                           'ListIndex = 0
    TrailerLength.AddItem "122"                          'ListIndex = 1
    TrailerLength.AddItem "146"                          'ListIndex = 2
    TrailerLength.AddItem "170"                          'ListIndex = 3
    TrailerLength.AddItem "194"                          'ListIndex = 4
    TrailerLength.AddItem "218"                          'ListIndex = 5
    TrailerLength.AddItem "242"                          'ListIndex = 6
    TrailerLength.AddItem "266"                          'ListIndex = 7
    TrailerLength.AddItem "290"                          'ListIndex = 8
    TrailerLength.AddItem "314"                          'ListIndex = 9
    TrailerLength.AddItem "338"                          'ListIndex = 10
    TrailerLength.AddItem "362"                          'ListIndex = 11
    TrailerLength.AddItem "386"                          'ListIndex = 12
    TrailerLength.AddItem "410"                          'ListIndex = 13

    TongueLength.AddItem "39"                            'ListIndex = 0
    TongueLength.AddItem "42"                            'ListIndex = 1
    TongueLength.AddItem "47"                            'ListIndex = 2
    
    OCDist.AddItem "16"                                  'ListIndex = 0
    OCDist.AddItem "24"                                  'ListIndex = 1
    
    TrailerHeight.AddItem "66"                           'ListIndex = 0
    TrailerHeight.AddItem "72"                           'ListIndex = 1
    TrailerHeight.AddItem "78"                           'ListIndex = 2
    'TrailerHeight.AddItem "84"                           'ListIndex = 3
    
    'AxleLocation.AddItem ".50"                          'ListIndex = 0
    'AxleLocation.AddItem ".55"                          'ListIndex = 1
    AxleLocation.AddItem ".60"                           'ListIndex = 2
    'AxleLocation.AddItem ".65"                          'ListIndex = 3
    
    FrameMember.AddItem "Tube 2.0 x 3.0 x 11ga"          'ListIndex = 0
    FrameMember.AddItem "Tube 2.0 x 4.0 x 11ga"         'ListIndex = 1
    FrameMember.AddItem "Tube 2.0 x 5.0 x 11ga"         'ListIndex = 2
    FrameMember.AddItem "Tube 2.0 x 8.0 x 11ga"         'ListIndex = 3
    
    SideWallMember.AddItem "Z-CHANNEL 1.125 x 1.125"     'ListIndex = 0
    'SideWallMember.AddItem "Tube 1.0 x 1.5 x 16ga"      'ListIndex = 1
    'SideWallMember.AddItem "Tube 2.0 x 2.0 x 11ga"      'ListIndex = 2
    
    NoseStructure.AddItem "Flat Nose"                    'ListIndex = 0
    NoseStructure.AddItem "2ft Wedge"                     'ListIndex = 1
    'NoseStructure.AddItem "4ft Wedge"                     'ListIndex = 2
    'NoseStructure.AddItem "6ft Wedge"                     'ListIndex = 3
    
    TrailerType.AddItem "Motiv Raven Lite Cargo (Steel)"    'ListIndex = 0
    TrailerType.AddItem "Motiv Raven Lite Cargo (Aluminum)" 'ListIndex = 1
    TrailerType.AddItem "Motiv MSX Series"                  'ListIndex = 2
    'TrailerType.AddItem "Motiv RSX Gooseneck Car Hauler"   'ListIndex = 3
    'TrailerType.AddItem "Motiv RSX Car Hauler"             'ListIndex = 4
    'TrailerType.AddItem "Motiv SSX Series Stacker"         'ListIndex = 5
    'TrailerType.AddItem "Motiv Steel Snow Trailer"         'ListIndex = 6
    
    NoOfAxles.AddItem "Single Axle"                      'ListIndex = 0
    NoOfAxles.AddItem "Dual Axle"                        'ListIndex = 1
    NoOfAxles.AddItem "Triple Axle"                      'ListIndex = 2
    
    AxleRating.AddItem "3500 lb 0 Deg"                   'ListIndex = 0
    ' Raven only allowed to use 3500 lb axles
    AxleRating.AddItem "5200 lb 10 Deg up"              'ListIndex = 1
    AxleRating.AddItem "6000 lb 10 Deg up"              'ListIndex = 2
    AxleRating.AddItem "7000 lb 10 Deg up"              'ListIndex = 3
    AxleRating.AddItem "6000 lb 22.5 Deg up"            'ListIndex = 4
    AxleRating.AddItem "7000 lb 22.5 Deg up"            'ListIndex = 5
    AxleRating.AddItem "8000 lb 22.5 Deg up"            'ListIndex = 6
    AxleRating.AddItem "10000 lb 22.5 Deg up"           'ListIndex = 7
    
    AxleSpacing.AddItem "Standard Axle Spacing"           'ListIndex = 0
    AxleSpacing.AddItem "Spread Axle Spacing"             'ListIndex = 1
    
End Function

Function SideWall(RefPlane As String)

Dim OCD As Double
Dim TW As Double
Dim Weldment As Double
Dim LineCnt As Integer
Dim X As Integer
Dim i As Integer
' 110207 *************************************************************************************
Dim WOCD As Double
Dim Y As Double
Dim WeldmentType(100) As String
Dim SketchName As String
'**********************************************************************************************
Dim DoorTubeHeight As Double

Dim LAST_VERTICAL_POST As Double

Dim swApp As Object

Dim Part As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long
Dim RotationAngle(100) As Double

Set swApp = Application.SldWorks
Set Part = swApp.ActiveDoc

'# Setting the Input Parameters

TW = 1         ' Tube Width ????  ************************** Fix This **********************************
LineCnt = 0

WOCD = 24

DoorTubeHeight = 2

If RefPlane = "CS" Then
   boolstatus = Part.Extension.SelectByID2("CS Wall@" + CStr(PartNumber.Value) + ".SLDPRT", "CONFIGURATIONS", 0, 0, 0, False, 0, Nothing, 0)
   boolstatus = Part.ShowConfiguration2("CS Wall")
ElseIf RefPlane = "RS" Then
   boolstatus = Part.Extension.SelectByID2("RS Wall@" + CStr(PartNumber.Value) + ".SLDPRT", "CONFIGURATIONS", 0, 0, 0, False, 0, Nothing, 0)
   boolstatus = Part.ShowConfiguration2("RS Wall")
End If

'110207 *************************************************************************************
Part.ClearSelection2 True
Dim skSegment As Object
' *******************************************************************************************

' Select Working Plane
boolstatus = Part.Extension.SelectByID2(RefPlane, "PLANE", 0, 0, 0, False, 0, Nothing, 0)
Part.SketchManager.InsertSketch True

If TrailerType.Value = "Motiv Raven Lite Cargo (Steel)" Then ' ******************************** Motiv Raven Lite Cargo (Steel) ************************************

   ' ****** Side Wall Header ******
   P1x = FW_Offset
   P1y = TrailerTotalHeight - (SWHeaderTubeHeight / 2)
   P2x = TrailerLength.Value - 2#
   P2y = TrailerTotalHeight - (SWHeaderTubeHeight / 2)
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   LineCnt = LineCnt + 1
   'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"
   
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", RefPlane, WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
   
   ' *** Set Offset Option in Menu ***

   If DoorOption Then
   ' ****** Side Wall Footer ******
      P1x = FW_Offset
      P1y = 0 + (SWFooterTubeHeight / 2)
      P2x = 16 + FW_Offset
      P2y = 0 + (SWFooterTubeHeight / 2)
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", RefPlane, WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
      
      ' ****** Side Wall Footer ******
      P1x = 48 + FW_Offset
      P1y = 0 + (SWFooterTubeHeight / 2)
      P2x = TrailerLength.Value - 2#
      P2y = 0 + (SWFooterTubeHeight / 2)
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", RefPlane, WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
      
   Else
      ' ****** Side Wall Footer ******
      P1x = FW_Offset
      P1y = 0 + (SWFooterTubeHeight / 2)
      P2x = TrailerLength.Value - 2#
      P2y = 0 + (SWFooterTubeHeight / 2)
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", RefPlane, WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
      
   End If

   '# ****** Vertical Members ******
   '# Note: Repeat this Operation based on Length and Frame Center to Center Distance

   X = (TrailerLength.Value / WOCD)
   
   P1x = Y + FW_Offset + (SWFrontVerticalTubeHeight / 2)
   P1y = SWHeaderTubeHeight
   P2x = Y + FW_Offset + (SWFrontVerticalTubeHeight / 2)
   P2y = TrailerTotalHeight - SWHeaderTubeHeight
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   Y = (WOCD)
   Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   LineCnt = LineCnt + 1
   'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"
   
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", RefPlane, WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
      
   For i = 2 To X
      
      If DoorOption And i < 3 Then
      
         ' ****** Side Wall Door Front Vertical Member ******
         P1x = 16 + FW_Offset - DoorTubeHeight / 2
         P1y = SWHeaderTubeHeight
         P2x = 16 + FW_Offset - DoorTubeHeight / 2
         P2y = TrailerTotalHeight - 8.5
         ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
         LAST_VERTICAL_POST = Y + FW_Offset
               
         If RefPlane = "CS" Then
            Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
         ElseIf RefPlane = "RS" Then
            Set skSegment = Part.SketchManager.CreateLine(P2x, P2y, 0#, P1x, P1y, 0#)
         End If
         LineCnt = LineCnt + 1
         'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
         RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
         WeldmentType(LineCnt) = "Tube 1.0 x 2.0 x 14ga"
         
         Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
         Write #1, "WALL", RefPlane, WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
         
         ' ****** Side Wall Door Rear Vertical Member ******
         P1x = 48 + FW_Offset + DoorTubeHeight / 2
         P1y = SWHeaderTubeHeight
         P2x = 48 + FW_Offset + DoorTubeHeight / 2
         P2y = TrailerTotalHeight - SWHeaderTubeHeight
         ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
         LAST_VERTICAL_POST = Y + FW_Offset
       
         If RefPlane = "CS" Then
            Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
         ElseIf RefPlane = "RS" Then
            Set skSegment = Part.SketchManager.CreateLine(P2x, P2y, 0#, P1x, P1y, 0#)
         End If
         LineCnt = LineCnt + 1
         'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
         RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
         WeldmentType(LineCnt) = "Tube 1.0 x 2.0 x 14ga"
         
         Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
         Write #1, "WALL", RefPlane, WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
                  
         ' ****** Side Wall Door Vertical Member in middle of Door ******
         P1x = 24 + FW_Offset
         P1y = TrailerTotalHeight - 8.5 + (DoorTubeHeight)
         P2x = 24 + FW_Offset
         P2y = TrailerTotalHeight - SWHeaderTubeHeight
         ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
         LAST_VERTICAL_POST = Y + 0.035 + FW_Offset
         
      
         If RefPlane = "CS" Then
            Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
         ElseIf RefPlane = "RS" Then
            Set skSegment = Part.SketchManager.CreateLine(P2x, P2y, 0#, P1x, P1y, 0#)
         End If
         LineCnt = LineCnt + 1
         'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
         RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
         WeldmentType(LineCnt) = "Tube 1.0 x 2.0 x 14ga"
         
         Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
         Write #1, "WALL", RefPlane, WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
         
         ' ****** Side Wall Door Header ******
         P1x = FW_Offset + SWFrontVerticalTubeHeight
         P1y = TrailerTotalHeight - 2.5 - 6 + (DoorTubeHeight / 2)
         P2x = 48 + FW_Offset
         P2y = TrailerTotalHeight - 2.5 - 6 + (DoorTubeHeight / 2)
         ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
         Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
         LineCnt = LineCnt + 1
         'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
         RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
         WeldmentType(LineCnt) = "Tube 1.0 x 2.0 x 14ga"

         Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
         Write #1, "WALL", RefPlane, WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.

         i = i + 1
         Y = (i * WOCD)
      Else
         P1x = Y + 0.035 + FW_Offset
         P1y = SWHeaderTubeHeight
         P2x = Y + 0.035 + FW_Offset
         P2y = TrailerTotalHeight - SWHeaderTubeHeight
         ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
         LAST_VERTICAL_POST = Y + 0.035 + FW_Offset
         Y = (i * WOCD)
      
         If RefPlane = "CS" Then
            Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
         ElseIf RefPlane = "RS" Then
            Set skSegment = Part.SketchManager.CreateLine(P2x, P2y, 0#, P1x, P1y, 0#)
         End If
         LineCnt = LineCnt + 1
         'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
         RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
         WeldmentType(LineCnt) = "Z-CHANNEL 1.125 x 1.125"
         
         Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
         Write #1, "WALL", RefPlane, WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
         
      End If
      
   Next i

   P1x = TrailerLength.Value - 2# - (1.125)
   P1y = SWHeaderTubeHeight
   P2x = TrailerLength.Value - 2# - (1.125)
   P2y = TrailerTotalHeight - SWHeaderTubeHeight
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   Y = (i * WOCD)
   
   If RefPlane = "CS" Then
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   ElseIf RefPlane = "RS" Then
      Set skSegment = Part.SketchManager.CreateLine(P2x, P2y, 0#, P1x, P1y, 0#)
   End If
   LineCnt = LineCnt + 1
   'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = "Z-CHANNEL 1.125 x 1.125"
     
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", RefPlane, WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
   
   '  ******************************************** BACKER MATERIAL FOR RAVEN **********************************
   '  *** Rear Backer ***
   P1x = TrailerLength.Value - 2# - 1.5                  ' Allow for .375" Clearance on the Length
   P1y = 6
   P2x = LAST_VERTICAL_POST + 0.375                      ' TrailerLength.Value - 3.5 - 1.125 - (16)
   P2y = 6
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   LineCnt = LineCnt + 1
   'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = "PLATE 2.0 x 14ga"
   
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", RefPlane, WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
   
   ' *** Front Backer ***
   P1x = FW_Offset + (SWFrontVerticalTubeHeight)                             ' Flush to the Front Tube
   P1y = 6
   
   If DoorOption Then
      P2x = (SWFrontVerticalTubeHeight) + 12.5 + FW_Offset              ' TrailerLength.Value - 3.5 - 1.125 - (16)
   Else
      P2x = (SWFrontVerticalTubeHeight) + 22.125 + 0.035 + FW_Offset               ' TrailerLength.Value - 3.5 - 1.125 - (16)
   End If
   
   P2y = 6
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   LineCnt = LineCnt + 1
   'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = "PLATE 2.0 x 14ga"
   
   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", RefPlane, WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
   
   Dim cnt As Double
   Dim WL As Double
   
   If TrailerWidth.Value = 84 Then
   Else
      '  *** Backer for Wheel Well ***
      For cnt = 1 To X
         WL = FW_Offset + cnt * WOCD
         If WL > AxleCL Then
            P2x = cnt * WOCD + 0.035 + FW_Offset
            cnt = X
         Else
            P1x = cnt * WOCD + 0.035 + FW_Offset
         End If
      Next cnt
   
      If DoorOption And CInt(TrailerLength.Value) < 144 Then
         P1x = P1x + 2
      Else
         P1x = P1x + 0.25
      End If
   
   
      'P1x = P1x + 0.25                      ' Need to Know Axle Location
      P1y = 13
      P2x = P2x - 0.25               ' TrailerLength.Value - 3.5 - 1.125 - (16)
      P2y = 13
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = "PLATE 2.0 x 14ga"
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", RefPlane, WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
   End If
   
   
   'WeldmentType(I) = "PLATE 2.0 x 14ga"

ElseIf TrailerType.Value = "Motiv MSX Series" Then  ' ******************************** Motiv MSX Series ************************************
   
   Dim SWH_Height As Double ' Side Wall Header Tube Height
   Dim SWH_Width As Double  ' Side Wall Header Tube Width
   
 '# ************* Side Wall Header
   SWH_Height = 3 ' Side Wall Header Tube Height
   SWH_Width = 1# ' Side Wall Header Tube Width
   
   P1x = FW_Offset
   P1y = TrailerTotalHeight - (SWH_Height / 2)
   P2x = TrailerLength.Value
   P2y = TrailerTotalHeight - (SWH_Height / 2)
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)

   '# Create Sketch Line
   Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   LineCnt = LineCnt + 1
   'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = "Tube 1.0 x 3.0 x 14ga"

   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", RefPlane, WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.

   If TrailerWidth.Value < 84 Then
   
      Dim SWF_Height As Double ' Side Wall Footer Tube Height
      Dim SWF_Width As Double  ' Side Wall Footer Tube Width

      '# Side Wall Footer
      SWF_Height = 1.5 ' Side Wall Footer Tube Height
      SWF_Width = 1#   ' Side Wall Footer Tube Width
      P1x = FW_Offset
      P1y = 0 + (SWF_Height / 2)
      P2x = TrailerLength.Value
      P2y = 0 + (SWF_Height / 2)
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"
   
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", RefPlane, WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
   Else
 
      ReturnVal = SideWallTubeSize() ' Call for Side Wall Checks *********************Fix This for Global Variables ****************************************
      
      Dim WWH As Double ' Wheel Well Height

      ReturnVal = WheelWellHeight(WWH)
                  
      Dim FVT_Height As Double ' First Vertical Tube Tube Height
      Dim FVT_Width As Double  ' First Vertical Tube Tube Width

      ' ***** Front Wheel Well Vertical Tube
      FVT_Height = 3#  ' First Vertical Tube Tube Height
      FVT_Width = 1#   ' First Vertical Tube Tube Width
      P1x = AxleCL - SWAxleClearance / 2 - FVT_Height / 2
      P1y = SWF_Height
      P2x = AxleCL - SWAxleClearance / 2 - FVT_Height / 2
      P2y = WWH
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = "Tube 1.0 x 3.0 x 14ga"
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", RefPlane, WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
      
      ' ***** Rear Wheel Well Vertical Tube
      FVT_Height = 3#  ' First Vertical Tube Tube Height
      FVT_Width = 1#   ' First Vertical Tube Tube Width
      P1x = AxleCL + SWAxleClearance / 2 + FVT_Height / 2
      P1y = 0
      P2x = AxleCL + SWAxleClearance / 2 + FVT_Height / 2
      P2y = WWH
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = "Tube 1.0 x 3.0 x 14ga"
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", RefPlane, WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.

   
      'Dim SWF_Height As Double ' Side Wall Footer Tube Height
      'Dim SWF_Width As Double  ' Side Wall Footer Tube Width

      '# ***** Side Wall Footer Front Piece
      SWF_Height = 1.5 ' Side Wall Footer Tube Height
      SWF_Width = 1#   ' Side Wall Footer Tube Width
      P1x = FW_Offset
      P1y = 0 + (SWF_Height / 2)
      P2x = AxleCL - SWAxleClearance / 2 - FVT_Height
      P2y = 0 + (SWF_Height / 2)
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", RefPlane, WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
      
      '# ***** Side Wall Footer Rear Piece
      SWF_Height = 1.5 ' Side Wall Footer Tube Height
      SWF_Width = 1#   ' Side Wall Footer Tube Width
      P1x = AxleCL + SWAxleClearance / 2 + FVT_Height
      P1y = 0 + (SWF_Height / 2)
      P2x = TrailerLength.Value
      P2y = 0 + (SWF_Height / 2)
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", RefPlane, WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
      
      Dim SWFWW_Height As Double
      Dim SWFWW_Width As Double
      
      '# ***** Side Wall Footer Over Wheel Well
      SWFWW_Height = 3#  ' Side Wall Footer Tube Height
      SWFWW_Width = 1#   ' Side Wall Footer Tube Width
      P1x = AxleCL - SWAxleClearance / 2 - SWRailOverlap
      P1y = WWH + SWFWW_Height / 2
      P2x = AxleCL + SWAxleClearance / 2 + SWRailOverlap
      P2y = WWH + SWFWW_Height / 2
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = "Tube 1.0 x 3.0 x 14ga"

      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", RefPlane, WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
      
   End If ' **************** If TrailerWidth.Value < 84 Then ***************************

   'Dim FVT_Height As Double ' First Vertical Tube Tube Height
   'Dim FVT_Width As Double  ' First Vertical Tube Tube Width


   ' ***** First Vertical Tube
   FVT_Height = 3#  ' First Vertical Tube Tube Height
   FVT_Width = 1#   ' First Vertical Tube Tube Width
   P1x = FW_Offset + (FVT_Height / 2)
   P1y = SWF_Height
   P2x = FW_Offset + (FVT_Height / 2)
   P2y = TrailerTotalHeight - SWH_Height
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   LineCnt = LineCnt + 1
   'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = "Tube 1.0 x 3.0 x 14ga"

   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", RefPlane, WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.

   Dim LVT_Height As Double ' Last Vertical Tube Tube Height
   Dim LVT_Width As Double  ' Last Vertical Tube Tube Width

   ' ***** Last Vertical Tube
   LVT_Height = 1.5 ' Last Vertical Tube Tube Height
   LVT_Width = 1#   ' Last Vertical Tube Tube Width
   P1x = TrailerLength.Value - LVT_Height / 2
   P1y = SWF_Height
   P2x = TrailerLength.Value - LVT_Height / 2
   P2y = TrailerTotalHeight - SWH_Height
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   LineCnt = LineCnt + 1
   'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"

   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", RefPlane, WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.

   ' Add Vertical Side Wall Members in Series

   ' ***** Second Vertical Tube @ 13 inches
   FVT_Height = 1.5  ' First Vertical Tube Tube Height
   FVT_Width = 1#   ' First Vertical Tube Tube Width
   P1x = FW_Offset + (FVT_Height / 2) + 13
   P1y = SWF_Height
   P2x = FW_Offset + (FVT_Height / 2) + 13
   P2y = TrailerTotalHeight - SWH_Height
   ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
   Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
   LineCnt = LineCnt + 1
   'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
   RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
   WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"

   Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
   Write #1, "WALL", RefPlane, WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.

   WOCD = 16
   X = ((TrailerLength.Value - 5 - 13 - 8) / WOCD)
   For i = 1 To X
      ' ***** Second Vertical Tube @ 13 inches
      FVT_Height = 1.5  ' First Vertical Tube Tube Height
      FVT_Width = 1#   ' First Vertical Tube Tube Width
      P1x = FW_Offset + (FVT_Height / 2) + 13 + i * 16
      P1y = SWF_Height
      P2x = FW_Offset + (FVT_Height / 2) + 13 + i * 16
      P2y = TrailerTotalHeight - SWH_Height
      ReturnVal = Convert2Meters(P1x, P1y, P2x, P2y)
      Set skSegment = Part.SketchManager.CreateLine(P1x, P1y, 0#, P2x, P2y, 0#)
      LineCnt = LineCnt + 1
      'RotationAngle(LineCnt) = 1.570796326795          ' 90 Degree Rotation Angle
      RotationAngle(LineCnt) = 0                        '  0 Degree Rotation Angle
      WeldmentType(LineCnt) = "Tube 1.0 x 1.5 x 16ga"
      
      Dist = (Sqr((P1x - P2x) ^ 2 + (P1y - P2y) ^ 2)) * 39.37
      Write #1, "WALL", RefPlane, WeldmentType(LineCnt), Dist, 0, 0  ' Write comma-delimited data.
      
   Next i

ElseIf TrailerType.Value = "Motiv RSX Gooseneck Car Hauler" Then
   MsgBox ("Not Yet")
   End
ElseIf TrailerType.Value = "Motiv RSX Car Hauler" Then
   MsgBox ("Not Yet")
   End
ElseIf TrailerType.Value = "Motiv SSX Series Stacker" Then
   MsgBox ("Not Yet")
   End
ElseIf TrailerType.Value = "Motiv Steel Snow Trailer" Then
   MsgBox ("Not Yet")
   End

End If

SketchName = Part.SketchManager.ActiveSketch.Name

'110207 **********************************************************************
Dim Feature As Object

Set swApp = Application.SldWorks

Set Part = swApp.ActiveDoc
Part.ViewZoomtofit2
Part.ClearSelection2 True
boolstatus = Part.Extension.SelectByID2("Point1@Origin", "EXTSKETCHPOINT", 0, 0, 0, True, 6, Nothing, 0)
longstatus = Part.SketchManager.FullyDefineSketch(1, 1, 1023, 1, 1, Nothing, -1, Nothing, 1, -1)
Part.ClearSelection2 True
boolstatus = Part.EditRebuild3

' ****************************************************************************
' Add Variable for Sketch Plane
Dim SketchPlane As String

SketchPlane = RefPlane + "_SK" ' Specify the SketchName

' Call Function to Create Weldments Based on Populated Variables Above
ReturnVal = BuildWeldment(SketchName, SketchPlane, WeldmentType(), LineCnt, RotationAngle())

End Function

Function RavenWedge(i As Integer)
   
   Dim swApp As Object
   Dim Part As Object
   Dim SelMgr As Object
   Dim boolstatus As Boolean
   Dim longstatus As Long, longwarnings As Long
   Dim Feature As Object
   Dim skSegment As Object
   Dim SketchName As String
   Dim SketchPlane As String
   
   Set swApp = Application.SldWorks

   Set Part = swApp.ActiveDoc
   Set SelMgr = Part.SelectionManager
   
   ' Select Working Plane
      boolstatus = Part.Extension.SelectByID2("FWR" + CStr(i), "PLANE", 0, 0, 0, False, 0, Nothing, 0)
      Part.SketchManager.InsertSketch True
    
   If TrailerWidth.Value = 60 Then

      Part.ClearSelection2 True
      Set skSegment = Part.SketchManager.CreateArc(-0.4572, 0, 0#, -0.7493, 0, 0#, -0.619228, 0.243042, 0#, -1)
      Part.ClearSelection2 True
      Set skSegment = Part.SketchManager.CreateArc(0#, 0.3048, 0#, -0.162028, 0.547842, 0#, 0.162028, 0.547842, 0#, -1)
      Part.ClearSelection2 True
      Set skSegment = Part.SketchManager.CreateArc(0.4572, 0#, 0#, 0.619228, 0.243042, 0#, 0.7493, 0#, 0#, -1)
      Part.ClearSelection2 True
      Set skSegment = Part.SketchManager.CreateLine(-0.619228, 0.243042, 0#, -0.162028, 0.547842, 0#)
      Part.ClearSelection2 True
      Set skSegment = Part.SketchManager.CreateLine(0.162028, 0.547842, 0#, 0.619228, 0.243042, 0#)
      Part.ClearSelection2 True
      boolstatus = Part.Extension.SelectByID2("Point10", "SKETCHPOINT", -0.619228, 0.243042, 0#, False, 0, Nothing, 0)
      boolstatus = Part.Extension.SelectByID2("Point3", "SKETCHPOINT", -0.619228, 0.243042, 0#, True, 0, Nothing, 0)
      Part.SketchAddConstraints "sgMERGEPOINTS"
      Part.ClearSelection2 True
      boolstatus = Part.Extension.SelectByID2("Point11", "SKETCHPOINT", -0.162028, 0.547842, 0#, False, 0, Nothing, 0)
      boolstatus = Part.Extension.SelectByID2("Point4", "SKETCHPOINT", -0.162028, 0.547842, 0, True, 0, Nothing, 0)
      Part.SketchAddConstraints "sgMERGEPOINTS"
      Part.ClearSelection2 True
      boolstatus = Part.Extension.SelectByID2("Point12", "SKETCHPOINT", 0.162028, 0.547842, 0#, False, 0, Nothing, 0)
      boolstatus = Part.Extension.SelectByID2("Point6", "SKETCHPOINT", 0.162028, 0.547842, 0#, True, 0, Nothing, 0)
      Part.SketchAddConstraints "sgMERGEPOINTS"
      boolstatus = Part.Extension.SelectByID2("Point13", "SKETCHPOINT", 0.619228, 0.243042, 0#, False, 0, Nothing, 0)
      boolstatus = Part.Extension.SelectByID2("Point8", "SKETCHPOINT", 0.619228, 0.243042, 0#, True, 0, Nothing, 0)
      Part.SketchAddConstraints "sgMERGEPOINTS"
      Part.ClearSelection2 True
      
   ElseIf TrailerWidth.Value = 72 Then
      
      Part.ClearSelection2 True
      Set skSegment = Part.SketchManager.CreateArc(-0.6096, 0, 0#, -0.9017, 0, 0#, -0.740231, 0.261262, 0#, -1)
      Part.ClearSelection2 True
      Set skSegment = Part.SketchManager.CreateArc(0#, 0.3048, 0#, -0.130631, 0.566062, 0#, 0.130631, 0.566062, 0#, -1)
      Part.ClearSelection2 True
      Set skSegment = Part.SketchManager.CreateArc(0.6096, 0#, 0#, 0.740231, 0.261262, 0#, 0.9017, 0#, 0#, -1)
      Part.ClearSelection2 True
      Set skSegment = Part.SketchManager.CreateLine(-0.740231, 0.261262, 0#, -0.130631, 0.566062, 0#)
      Part.ClearSelection2 True
      Set skSegment = Part.SketchManager.CreateLine(0.130631, 0.566062, 0#, 0.740231, 0.261262, 0#)
      Part.ClearSelection2 True
      boolstatus = Part.Extension.SelectByID2("Point10", "SKETCHPOINT", -0.740231, 0.261262, 0#, False, 0, Nothing, 0)
      boolstatus = Part.Extension.SelectByID2("Point3", "SKETCHPOINT", -0.740231, 0.261262, 0#, True, 0, Nothing, 0)
      Part.SketchAddConstraints "sgMERGEPOINTS"
      Part.ClearSelection2 True
      boolstatus = Part.Extension.SelectByID2("Point11", "SKETCHPOINT", -0.130631, 0.566062, 0#, False, 0, Nothing, 0)
      boolstatus = Part.Extension.SelectByID2("Point4", "SKETCHPOINT", -0.130631, 0.566062, 0, True, 0, Nothing, 0)
      Part.SketchAddConstraints "sgMERGEPOINTS"
      Part.ClearSelection2 True
      boolstatus = Part.Extension.SelectByID2("Point12", "SKETCHPOINT", 0.130631, 0.566062, 0#, False, 0, Nothing, 0)
      boolstatus = Part.Extension.SelectByID2("Point6", "SKETCHPOINT", 0.130631, 0.566062, 0#, True, 0, Nothing, 0)
      Part.SketchAddConstraints "sgMERGEPOINTS"
      boolstatus = Part.Extension.SelectByID2("Point13", "SKETCHPOINT", 0.740231, 0.261262, 0#, False, 0, Nothing, 0)
      boolstatus = Part.Extension.SelectByID2("Point8", "SKETCHPOINT", 0.740231, 0.261262, 0#, True, 0, Nothing, 0)
      Part.SketchAddConstraints "sgMERGEPOINTS"
      Part.ClearSelection2 True
      
   ElseIf TrailerWidth.Value = 84 Then
      
      Part.ClearSelection2 True
      Set skSegment = Part.SketchManager.CreateArc(-0.762, 0, 0#, -1.0541, 0, 0#, -0.870434, 0.271388, 0#, -1)
      Part.ClearSelection2 True
      Set skSegment = Part.SketchManager.CreateArc(0, 0.3048, 0#, -0.108434, 0.576028, 0#, 0.108434, 0.576028, 0#, -1)
      Part.ClearSelection2 True
      Set skSegment = Part.SketchManager.CreateArc(0.762, 0, 0#, 0.870434, 0.271388, 0, 1.0541, 0, 0#, -1)
      Part.ClearSelection2 True
      Set skSegment = Part.SketchManager.CreateLine(-0.870434, 0.271388, 0#, -0.108434, 0.576028, 0#)
      Part.ClearSelection2 True
      Set skSegment = Part.SketchManager.CreateLine(0.108434, 0.576028, 0#, 0.870434, 0.271388, 0#)
      Part.ClearSelection2 True
      'boolstatus = Part.Extension.SelectByID2("Point10", "SKETCHPOINT", -0.870434, 0.271388, 0#, False, 0, Nothing, 0)
      'boolstatus = Part.Extension.SelectByID2("Point3", "SKETCHPOINT", -0.870434, 0.271388, 0#, True, 0, Nothing, 0)
      'Part.SketchAddConstraints "sgMERGEPOINTS"
      boolstatus = Part.Extension.SelectByID2("Point2", "SKETCHPOINT", -0.870378778971, 0.271249793122, 0, False, 0, Nothing, 0)
      boolstatus = Part.Extension.SelectByID2("Point10", "SKETCHPOINT", -0.870434, 0.271388, 0, True, 0, Nothing, 0)
      Part.SketchAddConstraints "sgMERGEPOINTS"
      Part.ClearSelection2 True
      boolstatus = Part.Extension.SelectByID2("Point11", "SKETCHPOINT", -0.108434, 0.576028, 0#, False, 0, Nothing, 0)
      boolstatus = Part.Extension.SelectByID2("Point4", "SKETCHPOINT", -0.108434, 0.576028, 0, True, 0, Nothing, 0)
      Part.SketchAddConstraints "sgMERGEPOINTS"
      Part.ClearSelection2 True
      boolstatus = Part.Extension.SelectByID2("Point12", "SKETCHPOINT", 0.108434, 0.576028, 0#, False, 0, Nothing, 0)
      boolstatus = Part.Extension.SelectByID2("Point6", "SKETCHPOINT", 0.108434, 0.576028, 0#, True, 0, Nothing, 0)
      Part.SketchAddConstraints "sgMERGEPOINTS"
      boolstatus = Part.Extension.SelectByID2("Point13", "SKETCHPOINT", 0.870434, 0.271388, 0#, False, 0, Nothing, 0)
      boolstatus = Part.Extension.SelectByID2("Point8", "SKETCHPOINT", 0.870434, 0.271388, 0#, True, 0, Nothing, 0)
      Part.SketchAddConstraints "sgMERGEPOINTS"
      Part.ClearSelection2 True
      
   End If
      
      SketchName = Part.SketchManager.ActiveSketch.Name
      boolstatus = Part.EditRebuild3()

      SketchPlane = "FW_SK" + CStr(i) ' Specify the SketchName
      
      '110207 Define Sketch Plane To Add Weldments ***********************************************
      boolstatus = Part.Extension.SelectByID2(SketchName, "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
      boolstatus = Part.Extension.SelectByID2(SketchName, "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
      boolstatus = Part.SelectedFeatureProperties(0, 0, 0, 0, 0, 0, 0, 1, 0, SketchPlane)
      '*******************************************************************************************
            
      boolstatus = Part.Extension.SelectByID2("Arc1" + "@" + SketchPlane, "EXTSKETCHSEGMENT", -0.8135217069273, 0.2147028247387, 0, True, 0, Nothing, 0)
      boolstatus = Part.Extension.SelectByID2("Line1" + "@" + SketchPlane, "EXTSKETCHSEGMENT", -0.5612724783935, 0.5561735900352, 0, True, 0, Nothing, 0)
      boolstatus = Part.Extension.SelectByID2("Arc2" + "@" + SketchPlane, "EXTSKETCHSEGMENT", -0.2146892634799, 0.7760249576369, 0, True, 0, Nothing, 0)
      boolstatus = Part.Extension.SelectByID2("Line2" + "@" + SketchPlane, "EXTSKETCHSEGMENT", 0.2859384704661, 0.6694395559238, 0, True, 0, Nothing, 0)
      boolstatus = Part.Extension.SelectByID2("Arc3" + "@" + SketchPlane, "EXTSKETCHSEGMENT", 0.7520901492298, 0.3233413687194, 0, True, 0, Nothing, 0)
      Dim myFeature As Object
      Dim vGroups As Variant
      Dim GroupArray() As Object
      ReDim GroupArray(0 To 0) As Object
      Dim Group1 As Object
      Set Group1 = Part.FeatureManager.CreateStructuralMemberGroup()
      Dim vSegement1 As Variant
      Dim SegementArray1() As Object
      ReDim SegementArray1(0 To 4) As Object
      Part.ClearSelection2 True
      boolstatus = Part.Extension.SelectByID2("Arc1" + "@" + SketchPlane, "EXTSKETCHSEGMENT", -1.894998711036, 2.585805199894, 0, True, 0, Nothing, 0)
      boolstatus = Part.Extension.SelectByID2("Line1" + "@" + SketchPlane, "EXTSKETCHSEGMENT", -1.894998711036, 2.585805199894, 0, True, 0, Nothing, 0)
      boolstatus = Part.Extension.SelectByID2("Arc2" + "@" + SketchPlane, "EXTSKETCHSEGMENT", -1.894998711036, 2.585805199894, 0, True, 0, Nothing, 0)
      boolstatus = Part.Extension.SelectByID2("Line2" + "@" + SketchPlane, "EXTSKETCHSEGMENT", -1.894998711036, 2.585805199894, 0, True, 0, Nothing, 0)
      boolstatus = Part.Extension.SelectByID2("Arc3" + "@" + SketchPlane, "EXTSKETCHSEGMENT", -1.894998711036, 2.585805199894, 0, True, 0, Nothing, 0)
      Dim Segment As Object
      Set Segment = Part.SelectionManager.GetSelectedObject5(1)
      Set SegementArray1(0) = Segment
      Set Segment = Part.SelectionManager.GetSelectedObject5(2)
      Set SegementArray1(1) = Segment
      Set Segment = Part.SelectionManager.GetSelectedObject5(3)
      Set SegementArray1(2) = Segment
      Set Segment = Part.SelectionManager.GetSelectedObject5(4)
      Set SegementArray1(3) = Segment
      Set Segment = Part.SelectionManager.GetSelectedObject5(5)
      Set SegementArray1(4) = Segment
      vSegement1 = SegementArray1
      Group1.Segments = (vSegement1)
      Group1.ApplyCornerTreatment = True
      Group1.CornerTreatmentType = 1
      Group1.GapWithinGroup = 0
      Group1.GapForOtherGroups = 0
      Group1.Angle = 0
      Set GroupArray(0) = Group1
      vGroups = GroupArray
      Set myFeature = Part.FeatureManager.InsertStructuralWeldment4(WeldmentPath + "TUBES\1 x 1 x 16ga.sldlfp", 1, True, (vGroups))
      Part.ClearSelection2 True
        
      
End Function

Function Convert2Meters(P1x As Double, P1y As Double, P2x As Double, P2y As Double)
   P1x = P1x * 0.0254
   P1y = P1y * 0.0254
   P2x = P2x * 0.0254
   P2y = P2y * 0.0254
End Function
