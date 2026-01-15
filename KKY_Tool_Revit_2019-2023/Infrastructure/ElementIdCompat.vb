Imports Autodesk.Revit.DB

Namespace Infrastructure

    ''' <summary>
    ''' ElementId 관련 버전 호환 헬퍼
    ''' - 2019~2023: IntegerValue / New ElementId(Integer)
    ''' - 2025: Value / New ElementId(Long)
    ''' </summary>
    Friend Module ElementIdCompat

        ''' <summary>
        ''' ElementId를 Int32로 꺼내는 공용 함수
        ''' </summary>
        <Runtime.CompilerServices.Extension>
        Public Function IntValue(id As ElementId) As Integer
            If id Is Nothing Then Return -1
#If REVIT2025 Then
            ' Revit 2024+ 권장: Value (Int64) 사용
            Return CInt(id.Value)
#Else
            ' Revit 2019~2023: 기존 IntegerValue
            Return id.IntegerValue
#End If
        End Function

        ''' <summary>
        ''' Int32에서 ElementId 생성 (버전별 생성자 호환)
        ''' </summary>
        Public Function FromInt(id As Integer) As ElementId
#If REVIT2025 Then
            ' 2024+ 권장: Int64 생성자 사용
            Return New ElementId(CLng(id))
#Else
            ' 2019~2023: 기존 Int32 생성자
            Return New ElementId(id)
#End If
        End Function

        ''' <summary>
        ''' Int64에서 ElementId 생성 (모든 버전 공통)
        ''' </summary>
        Public Function FromLong(id As Long) As ElementId
#If REVIT2025 Then
            ' 2024+ : Int64 생성자 사용
            Return New ElementId(id)
#Else
            ' 2019~2023: Int32 생성자만 존재
            Return New ElementId(CInt(id))
#End If
        End Function

    End Module

End Namespace
