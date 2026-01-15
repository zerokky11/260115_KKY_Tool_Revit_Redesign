Option Explicit On
Option Strict On

Imports Autodesk.Revit.DB
Imports System.Runtime.CompilerServices

' Revit 2019~2023에서 사용하던 BuiltInParameterGroup 을
' Revit 2025 (GroupTypeId 기반)에서만 호환용으로 제공하는 쉼 레이어
' ▶ Global.Autodesk.Revit.DB 로 선언해서
'    실제 타입 이름이 정확히 Autodesk.Revit.DB.BuiltInParameterGroup 이 되도록 한다.
Namespace Global.Autodesk.Revit.DB

    ''' <summary>
    ''' 최소 호환용 BuiltInParameterGroup enum.
    ''' 실제 값은 중요하지 않고, GroupTypeId 매핑에만 사용된다.
    ''' ParamPropagateService.vb 에서 사용하는 그룹들만 넣어두고,
    ''' 나머지는 Else 분기로 Data 그룹으로 처리한다.
    ''' </summary>
    Public Enum BuiltInParameterGroup
        PG_DATA
        PG_TEXT
        PG_GEOMETRY
        PG_CONSTRAINTS
        PG_IDENTITY_DATA
        PG_MATERIALS
        PG_GRAPHICS
        PG_ANALYSIS
        PG_GENERAL
    End Enum

    ''' <summary>
    ''' FamilyManager.AddParameter/ReplaceParameter 에
    ''' BuiltInParameterGroup 을 그대로 넘길 수 있도록 하는 확장 메서드.
    ''' 내부에서는 Revit 2025 API가 요구하는 GroupTypeId로 변환한다.
    ''' </summary>
    Public Module BuiltInParameterGroupCompat

        <Extension>
        Public Function ToGroupTypeId(group As BuiltInParameterGroup) As ForgeTypeId
            Return MapGroup(group)
        End Function

        Public Function FromGroupTypeId(groupId As ForgeTypeId) As BuiltInParameterGroup
            Return ToBuiltInGroup(groupId)
        End Function

        <Extension>
        Public Function IsInGroup(def As Definition, group As BuiltInParameterGroup) As Boolean
            If def Is Nothing Then Return False

            Return ToBuiltInGroup(def.GetGroupTypeId()) = group
        End Function

        Private Function MapGroup(group As BuiltInParameterGroup) As ForgeTypeId
            Select Case group
                Case BuiltInParameterGroup.PG_TEXT
                    Return GroupTypeId.Text
                Case BuiltInParameterGroup.PG_GEOMETRY
                    Return GroupTypeId.Geometry
                Case BuiltInParameterGroup.PG_CONSTRAINTS
                    Return GroupTypeId.Constraints
                Case BuiltInParameterGroup.PG_MATERIALS
                    Return GroupTypeId.Materials
                Case BuiltInParameterGroup.PG_GRAPHICS
                    Return GroupTypeId.Graphics

                ' 나머지 그룹들은 전부 Data 그룹으로 통일
                Case BuiltInParameterGroup.PG_DATA,
                     BuiltInParameterGroup.PG_IDENTITY_DATA,
                     BuiltInParameterGroup.PG_ANALYSIS,
                     BuiltInParameterGroup.PG_GENERAL
                    Return GroupTypeId.Data

                Case Else
                    Return GroupTypeId.Data
            End Select
        End Function

        Public Function ToBuiltInGroup(groupId As ForgeTypeId) As BuiltInParameterGroup
            If groupId Is Nothing Then
                Return BuiltInParameterGroup.PG_DATA
            End If

            If groupId.Equals(GroupTypeId.Text) Then
                Return BuiltInParameterGroup.PG_TEXT
            End If
            If groupId.Equals(GroupTypeId.Geometry) Then
                Return BuiltInParameterGroup.PG_GEOMETRY
            End If
            If groupId.Equals(GroupTypeId.Constraints) Then
                Return BuiltInParameterGroup.PG_CONSTRAINTS
            End If
            If groupId.Equals(GroupTypeId.Materials) Then
                Return BuiltInParameterGroup.PG_MATERIALS
            End If
            If groupId.Equals(GroupTypeId.Graphics) Then
                Return BuiltInParameterGroup.PG_GRAPHICS
            End If

            ' 나머지는 전부 Data로 처리
            Return BuiltInParameterGroup.PG_DATA
        End Function
    End Module

    Public Module FamilyManagerCompatExtensions

        <Extension>
        Public Function AddParameter(fm As FamilyManager,
                                     def As ExternalDefinition,
                                     group As BuiltInParameterGroup,
                                     isInstance As Boolean) As FamilyParameter

            Dim gId As ForgeTypeId = BuiltInParameterGroupCompat.ToGroupTypeId(group)
            Return fm.AddParameter(def, gId, isInstance)
        End Function

        <Extension>
        Public Function ReplaceParameter(fm As FamilyManager,
                                         param As FamilyParameter,
                                         def As ExternalDefinition,
                                         group As BuiltInParameterGroup,
                                         isInstance As Boolean) As FamilyParameter

            Dim gId As ForgeTypeId = BuiltInParameterGroupCompat.ToGroupTypeId(group)
            Return fm.ReplaceParameter(param, def, gId, isInstance)
        End Function

    End Module

End Namespace
