' Cmd_BatchAddSharedParameter.vb
' Revit 2019 / .NET Framework 4.8
' Prototype:
' - Shared Parameter source: current Revit Application.SharedParametersFilename
' - Select 1+ shared parameters, add to "Selected list"
' - Each parameter has its own settings dialog
' - Categories UI: TreeView with sub-categories (+) like Revit UI
' - FIX(2026-01): show non-bindable leaf categories (e.g., Stairs\Supports), allow selecting them,
'                 attempt binding anyway, and post-check log CAT_DROPPED_BY_REVIT.
' - Workshared: open with CloseAllWorksets, then SynchronizeWithCentral with comment
' - Non-workshared: Save()

Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Drawing

Imports WinForms = System.Windows.Forms

Imports Autodesk.Revit.Attributes
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports RevitApp = Autodesk.Revit.ApplicationServices.Application

Namespace Global.KKY_Tool_Revit

    <Transaction(TransactionMode.Manual)>
    Public Class Cmd_BatchAddSharedParameter
        Implements IExternalCommand

        Public Function Execute(commandData As ExternalCommandData,
                                ByRef message As String,
                                elements As ElementSet) As Result Implements IExternalCommand.Execute

            Dim uiapp As UIApplication = commandData.Application
            Dim app As RevitApp = uiapp.Application

            Dim baseDoc As Document = Nothing
            If uiapp.ActiveUIDocument IsNot Nothing Then
                baseDoc = uiapp.ActiveUIDocument.Document
            End If

            If baseDoc Is Nothing OrElse baseDoc.IsFamilyDocument Then
                TaskDialog.Show("Shared Parameter Batch", "활성 프로젝트 문서가 필요합니다. (패밀리 문서 불가)")
                Return Result.Cancelled
            End If

            ' Shared Parameter file must be set in Revit (Manage > Shared Parameters)
            Dim spFilePath As String = app.SharedParametersFilename
            If String.IsNullOrWhiteSpace(spFilePath) OrElse Not File.Exists(spFilePath) Then
                TaskDialog.Show("Shared Parameter Batch",
                                "현재 Revit에 설정된 Shared Parameter TXT를 찾을 수 없습니다." & vbCrLf &
                                "Revit에서 먼저 설정하세요: Manage > Shared Parameters" & vbCrLf & vbCrLf &
                                "현재 값: " & If(spFilePath, "(empty)"))
                Return Result.Cancelled
            End If

            ' Category tree (with Path; include non-bindable leaves too)
            Dim categoryTreeRoots As List(Of CategoryTreeItem) = BuildCategoryTree(baseDoc)
            If categoryTreeRoots.Count = 0 Then
                TaskDialog.Show("Shared Parameter Batch", "카테고리를 찾지 못했습니다.")
                Return Result.Cancelled
            End If

            Dim settings As BatchSettings = Nothing
            Using frm As New SharedParamBatchMainForm(app, spFilePath, categoryTreeRoots)
                Dim dr As WinForms.DialogResult = frm.ShowDialog()
                If dr <> WinForms.DialogResult.OK Then
                    Return Result.Cancelled
                End If
                settings = frm.GetSettings()
            End Using

            If settings Is Nothing Then Return Result.Cancelled

            Dim validationError As String = ValidateBatchSettings(settings, spFilePath)
            If validationError <> "" Then
                TaskDialog.Show("Shared Parameter Batch", validationError)
                Return Result.Cancelled
            End If

            Dim logs As New List(Of String)()
            Dim okCount As Integer = 0
            Dim failCount As Integer = 0
            Dim skipCount As Integer = 0
            Dim warnCount As Integer = 0

            Dim originalSpFile As String = app.SharedParametersFilename

            Dim progress As ProgressForm = Nothing

            Try
                progress = New ProgressForm()
                progress.Show()
                progress.SetStatus("Shared Parameter 정의 로드 중...")

                ' Ensure current shared parameter file
                app.SharedParametersFilename = spFilePath

                Dim defFile As DefinitionFile = app.OpenSharedParameterFile()
                If defFile Is Nothing Then
                    TaskDialog.Show("Shared Parameter Batch", "Shared Parameter 파일을 열 수 없습니다: " & spFilePath)
                    Return Result.Failed
                End If

                ' Build ExternalDefinition map by GUID for selected parameters
                Dim extDefByGuid As New Dictionary(Of Guid, ExternalDefinition)()
                For Each p As ParamToBind In settings.Parameters
                    Dim extDef As ExternalDefinition = FindExternalDefinitionByGuid(defFile, p.GuidValue)
                    If extDef Is Nothing Then
                        TaskDialog.Show("Shared Parameter Batch",
                                        "Shared Parameter 정의를 찾지 못했습니다." & vbCrLf &
                                        "- 파일: " & spFilePath & vbCrLf &
                                        "- 이름: " & p.ParamName & vbCrLf &
                                        "- GUID: " & p.GuidValue.ToString())
                        Return Result.Failed
                    End If

                    If Not extDefByGuid.ContainsKey(p.GuidValue) Then
                        extDefByGuid.Add(p.GuidValue, extDef)
                    End If
                Next

                Dim total As Integer = settings.RvtPaths.Count
                For i As Integer = 0 To total - 1
                    Dim rvtPath As String = settings.RvtPaths(i)
                    progress.SetProgress(i + 1, total)
                    progress.SetStatus("처리 중: " & rvtPath)
                    WinForms.Application.DoEvents()

                    If String.IsNullOrWhiteSpace(rvtPath) OrElse Not File.Exists(rvtPath) Then
                        skipCount += 1
                        logs.Add("[SKIP] 파일 없음: " & rvtPath)
                        Continue For
                    End If

                    If IsAlreadyOpen(app, rvtPath) Then
                        skipCount += 1
                        logs.Add("[SKIP] 이미 열려있음: " & rvtPath)
                        Continue For
                    End If

                    Dim doc As Document = Nothing
                    Try
                        doc = OpenProjectDocument(app, rvtPath, settings.CloseAllWorksetsOnOpen)

                        If doc Is Nothing Then
                            failCount += 1
                            logs.Add("[FAIL] 열기 실패(Unknown): " & rvtPath)
                            Continue For
                        End If

                        If doc.IsFamilyDocument Then
                            skipCount += 1
                            logs.Add("[SKIP] 패밀리 문서: " & rvtPath)
                            SafeClose(doc, False)
                            Continue For
                        End If

                        ' Apply bindings (multi parameters)
                        Dim perDocNotes As String = ""
                        Dim applied As Boolean = ApplyAllSharedParameterBindings(doc, app, extDefByGuid, settings.Parameters, perDocNotes)

                        If Not applied Then
                            failCount += 1
                            logs.Add("[FAIL] " & rvtPath & " :: " & perDocNotes)
                            SafeClose(doc, False)
                            Continue For
                        End If

                        If Not String.IsNullOrWhiteSpace(perDocNotes) Then
                            Dim lines As String() = perDocNotes.Split(New String() {vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
                            For Each ln As String In lines
                                warnCount += 1
                                logs.Add("[WARN] " & rvtPath & " :: " & ln)
                            Next
                        End If

                        ' Save / Sync
                        If doc.IsWorkshared Then
                            Dim syncLog As String = ""
                            Dim synced As Boolean = SyncWithCentral(doc, settings.SyncComment, syncLog)
                            If Not synced Then
                                failCount += 1
                                logs.Add("[FAIL] " & rvtPath & " :: Sync 실패 :: " & syncLog)
                                SafeClose(doc, False)
                                Continue For
                            End If

                            okCount += 1
                            logs.Add("[OK] " & rvtPath & " :: 적용(" & settings.Parameters.Count & "개) + Sync 완료 (Comment: " & settings.SyncComment & ")")
                            SafeClose(doc, False)
                        Else
                            doc.Save()
                            okCount += 1
                            logs.Add("[OK] " & rvtPath & " :: 적용(" & settings.Parameters.Count & "개) + Save 완료")
                            SafeClose(doc, False)
                        End If

                    Catch ex As Exception
                        failCount += 1
                        logs.Add("[FAIL] " & rvtPath & " :: 예외: " & ex.Message)
                        If doc IsNot Nothing Then
                            SafeClose(doc, False)
                        End If
                    End Try
                Next

            Catch exTop As Exception
                TaskDialog.Show("Shared Parameter Batch", "치명적 예외: " & exTop.Message)
                Return Result.Failed

            Finally
                Try
                    app.SharedParametersFilename = originalSpFile
                Catch
                End Try

                If progress IsNot Nothing Then
                    Try
                        progress.Close()
                    Catch
                    End Try
                End If
            End Try

            Dim logPath As String = Path.Combine(Path.GetTempPath(),
                                                "KKY_SharedParamBatch_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".txt")
            Try
                File.WriteAllLines(logPath, logs.ToArray(), Encoding.UTF8)
            Catch
            End Try

            Dim summary As String =
                "완료" & vbCrLf &
                "- OK: " & okCount & vbCrLf &
                "- FAIL: " & failCount & vbCrLf &
                "- SKIP: " & skipCount & vbCrLf &
                "- WARN: " & warnCount & vbCrLf & vbCrLf &
                "로그 파일:" & vbCrLf &
                logPath

            TaskDialog.Show("Shared Parameter Batch", summary)
            Return Result.Succeeded
        End Function

#Region "Data Models"

        Friend Class BatchSettings
            Public Property RvtPaths As List(Of String)
            Public Property Parameters As List(Of ParamToBind)
            Public Property CloseAllWorksetsOnOpen As Boolean
            Public Property SyncComment As String
        End Class

        Friend Class CategoryRef
            Public Property IdInt As Integer
            Public Property Path As String
            Public Property Name As String

            Public Function Clone() As CategoryRef
                Return New CategoryRef() With {.IdInt = Me.IdInt, .Path = Me.Path, .Name = Me.Name}
            End Function
        End Class

        Friend Class ParamBindingSettings
            Public Property IsInstanceBinding As Boolean
            Public Property ParamGroup As BuiltInParameterGroup
            Public Property AllowVaryBetweenGroups As Boolean

            ' store categories as (Id + Path)
            Public Property Categories As List(Of CategoryRef)

            Public Function CloneDeep() As ParamBindingSettings
                Dim c As New ParamBindingSettings()
                c.IsInstanceBinding = Me.IsInstanceBinding
                c.ParamGroup = Me.ParamGroup
                c.AllowVaryBetweenGroups = Me.AllowVaryBetweenGroups

                Dim src As List(Of CategoryRef) = If(Me.Categories, New List(Of CategoryRef)())
                Dim dst As New List(Of CategoryRef)()
                For Each x As CategoryRef In src
                    If x IsNot Nothing Then dst.Add(x.Clone())
                Next
                c.Categories = dst

                Return c
            End Function
        End Class

        Friend Class ParamToBind
            Public Property GroupName As String
            Public Property ParamName As String
            Public Property GuidValue As Guid
            Public Property ParamTypeLabel As String
            Public Property Description As String
            Public Property Settings As ParamBindingSettings

            Public ReadOnly Property GuidString As String
                Get
                    Return GuidValue.ToString()
                End Get
            End Property

            Public ReadOnly Property BindingDisplay As String
                Get
                    If Settings Is Nothing Then Return "-"
                    Return If(Settings.IsInstanceBinding, "Instance", "Type")
                End Get
            End Property

            Public ReadOnly Property ParamGroupDisplay As String
                Get
                    If Settings Is Nothing Then Return "-"
                    Try
                        Return LabelUtils.GetLabelFor(Settings.ParamGroup)
                    Catch
                        Return Settings.ParamGroup.ToString()
                    End Try
                End Get
            End Property

            Public ReadOnly Property CategoriesCountDisplay As String
                Get
                    If Settings Is Nothing OrElse Settings.Categories Is Nothing Then Return "0"
                    Return Settings.Categories.Count.ToString()
                End Get
            End Property
        End Class

        Friend Class CategoryTreeItem
            Public Property IdInt As Integer
            Public Property Name As String
            Public Property Path As String
            Public Property CatType As CategoryType
            Public Property IsBindable As Boolean
            Public Property Children As List(Of CategoryTreeItem)

            Public Sub New()
                Children = New List(Of CategoryTreeItem)()
            End Sub
        End Class

        Private Class CategoryMaps
            Public Sub New()
                ById = New Dictionary(Of Integer, Category)()
                ByPath = New Dictionary(Of String, Category)(StringComparer.OrdinalIgnoreCase)
            End Sub
            Public Property ById As Dictionary(Of Integer, Category)
            Public Property ByPath As Dictionary(Of String, Category)
        End Class

#End Region

#Region "Validation"

        Private Shared Function ValidateBatchSettings(s As BatchSettings, spFilePath As String) As String
            If s Is Nothing Then Return "설정이 비어있습니다."
            If String.IsNullOrWhiteSpace(spFilePath) OrElse Not File.Exists(spFilePath) Then
                Return "Shared Parameter TXT가 유효하지 않습니다."
            End If

            If s.RvtPaths Is Nothing OrElse s.RvtPaths.Count = 0 Then
                Return "대상 RVT 파일을 1개 이상 선택하세요."
            End If

            If s.Parameters Is Nothing OrElse s.Parameters.Count = 0 Then
                Return "추가할 Shared Parameter를 1개 이상 등록하세요."
            End If

            Dim noCat As New List(Of String)()
            For Each p As ParamToBind In s.Parameters
                If p Is Nothing OrElse p.Settings Is Nothing OrElse p.Settings.Categories Is Nothing OrElse p.Settings.Categories.Count = 0 Then
                    noCat.Add(If(p IsNot Nothing, p.ParamName, "(unknown)"))
                End If
            Next
            If noCat.Count > 0 Then
                Return "카테고리가 지정되지 않은 파라미터가 있습니다." & vbCrLf &
                       String.Join(", ", noCat.ToArray())
            End If

            Return ""
        End Function

#End Region

#Region "Shared Parameter Lookup"

        Private Shared Function FindExternalDefinitionByGuid(defFile As DefinitionFile, targetGuid As Guid) As ExternalDefinition
            If defFile Is Nothing Then Return Nothing

            For Each g As DefinitionGroup In defFile.Groups
                If g Is Nothing Then Continue For
                For Each d As Definition In g.Definitions
                    Dim ed As ExternalDefinition = TryCast(d, ExternalDefinition)
                    If ed Is Nothing Then Continue For
                    If ed.GUID = targetGuid Then Return ed
                Next
            Next

            Return Nothing
        End Function

#End Region

#Region "Apply Bindings (Multi-Parameter)"

        Private Shared Function ApplyAllSharedParameterBindings(doc As Document,
                                                               app As RevitApp,
                                                               extDefByGuid As Dictionary(Of Guid, ExternalDefinition),
                                                               paramList As List(Of ParamToBind),
                                                               ByRef notes As String) As Boolean
            notes = ""

            If doc Is Nothing OrElse app Is Nothing OrElse extDefByGuid Is Nothing OrElse paramList Is Nothing OrElse paramList.Count = 0 Then
                notes = "입력값 오류(doc/app/defs/params)."
                Return False
            End If

            Dim maps As CategoryMaps = BuildAvailableCategoryMaps(doc)

            Dim warnLines As New List(Of String)()

            Using tg As New TransactionGroup(doc, "Batch Add Shared Parameters")
                tg.Start()

                For Each p As ParamToBind In paramList
                    If p Is Nothing OrElse p.Settings Is Nothing Then
                        notes = "파라미터 설정이 비어있습니다."
                        tg.RollBack()
                        Return False
                    End If

                    Dim extDef As ExternalDefinition = Nothing
                    If Not extDefByGuid.TryGetValue(p.GuidValue, extDef) OrElse extDef Is Nothing Then
                        notes = "정의 누락: " & p.ParamName
                        tg.RollBack()
                        Return False
                    End If

                    Dim perParamWarn As String = ""
                    Dim perParamErr As String = ""

                    Dim ok As Boolean = ApplyOneSharedParameterBinding(doc, app, maps, extDef, p, perParamWarn, perParamErr)

                    If Not ok Then
                        notes = "Param [" & p.ParamName & "] :: " & perParamErr
                        tg.RollBack()
                        Return False
                    End If

                    If Not String.IsNullOrWhiteSpace(perParamWarn) Then
                        Dim lines As String() = perParamWarn.Split(New String() {vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
                        For Each ln As String In lines
                            warnLines.Add("Param [" & p.ParamName & "] :: " & ln)
                        Next
                    End If
                Next

                tg.Assimilate()
            End Using

            If warnLines.Count > 0 Then
                notes = String.Join(vbCrLf, warnLines.ToArray())
            Else
                notes = ""
            End If

            Return True
        End Function

        Private Shared Function ApplyOneSharedParameterBinding(doc As Document,
                                                              app As RevitApp,
                                                              maps As CategoryMaps,
                                                              extDef As ExternalDefinition,
                                                              p As ParamToBind,
                                                              ByRef warnLog As String,
                                                              ByRef errLog As String) As Boolean
            warnLog = ""
            errLog = ""

            If HasParameterNameConflict(doc, extDef) Then
                errLog = "동일한 파라미터 이름이 이미 문서에 존재합니다(다른 GUID 가능). 이름 충돌."
                Return False
            End If

            Dim req As List(Of CategoryRef) = If(p.Settings.Categories, New List(Of CategoryRef)())
            Dim warn As New List(Of String)()

            Dim catSetFull As CategorySet = app.Create.NewCategorySet()
            Dim catSetBindableOnly As CategorySet = app.Create.NewCategorySet()

            Dim attemptedFull As New Dictionary(Of Integer, CategoryRef)()
            Dim attemptedBindable As New Dictionary(Of Integer, CategoryRef)()

            Dim nonBindableSelected As Integer = 0

            For Each cref As CategoryRef In req
                If cref Is Nothing Then Continue For

                Dim resolveBy As String = ""
                Dim cat As Category = ResolveCategoryInDoc(maps, cref, resolveBy)

                If cat Is Nothing Then
                    warn.Add("CAT_NOT_FOUND: " & DescribeCategoryRef(cref))
                    Continue For
                End If

                Dim rid As Integer = cat.Id.IntegerValue

                ' read bindable flag (but DO NOT use it to hide/skip anymore)
                Dim bindableFlag As Boolean = False
                Try
                    bindableFlag = cat.AllowsBoundParameters
                Catch
                    bindableFlag = False
                End Try

                ' Always attempt insert into FULL set (so Supports can be attempted)
                If Not attemptedFull.ContainsKey(rid) Then
                    Try
                        catSetFull.Insert(cat)
                        attemptedFull.Add(rid, cref)
                    Catch exIns As Exception
                        warn.Add("CAT_INSERT_FAIL: " & DescribeCategoryRef(cref) & " :: " & exIns.Message)
                    End Try
                End If

                If bindableFlag Then
                    If Not attemptedBindable.ContainsKey(rid) Then
                        Try
                            catSetBindableOnly.Insert(cat)
                            attemptedBindable.Add(rid, cref)
                        Catch
                        End Try
                    End If
                Else
                    nonBindableSelected += 1
                    warn.Add("CAT_FLAG_NOT_BINDABLE(AllowsBoundParameters=False): " &
                             DescribeCategoryRef(cref) & " (ResolvedBy=" & resolveBy & ", ResolvedId=" & rid & ")")
                End If
            Next

            If attemptedFull.Count = 0 Then
                errLog = "선택된 카테고리가 이 문서에서 유효하지 않음(미존재/해석 실패/Insert 실패)."
                Return False
            End If

            Dim finalAttempted As Dictionary(Of Integer, CategoryRef) = attemptedFull
            Dim usedBindableOnly As Boolean = False

            Try
                Using t As New Autodesk.Revit.DB.Transaction(doc, "Bind Shared Parameter: " & p.ParamName)
                    t.Start()

                    Dim map As BindingMap = doc.ParameterBindings

                    Dim insertedOk As Boolean = False

                    ' 1) try FULL set
                    insertedOk = TryInsertOrReinsert(map, extDef, BuildBinding(app, catSetFull, p.Settings.IsInstanceBinding), p.Settings.ParamGroup)

                    ' 2) if it fails and we had non-bindable categories, fallback to bindable-only set
                    If Not insertedOk AndAlso nonBindableSelected > 0 AndAlso attemptedBindable.Count > 0 Then
                        warn.Add("BINDING_FALLBACK: full-set insert failed, retry bindable-only set (removed non-bindable candidates).")
                        insertedOk = TryInsertOrReinsert(map, extDef, BuildBinding(app, catSetBindableOnly, p.Settings.IsInstanceBinding), p.Settings.ParamGroup)

                        If insertedOk Then
                            finalAttempted = attemptedBindable
                            usedBindableOnly = True
                        End If
                    End If

                    If Not insertedOk Then
                        t.RollBack()
                        errLog = "ParameterBindings Insert/ReInsert 실패(이미 존재/제약/카테고리 문제)."
                        Return False
                    End If

                    t.Commit()
                End Using

                ' 안정화: 바인딩 후 regenerate (SharedParameterElement 생성 타이밍 보정)
                Try
                    doc.Regenerate()
                Catch
                End Try

                ' Vary between groups: apply AFTER binding commit
                If p.Settings IsNot Nothing AndAlso p.Settings.IsInstanceBinding Then
                    Dim vWarn As String = ""
                    Dim vOk As Boolean = TrySetAllowVaryBetweenGroups(doc, extDef.GUID, p.Settings.AllowVaryBetweenGroups, vWarn)
                    If Not vOk AndAlso Not String.IsNullOrWhiteSpace(vWarn) Then
                        warn.Add(vWarn)
                    End If
                End If

                ' Post-check: did Revit silently drop any categories?
                Dim boundIds As HashSet(Of Integer) = GetBoundCategoryIds(doc, extDef.GUID)
                If boundIds.Count > 0 Then
                    For Each kv As KeyValuePair(Of Integer, CategoryRef) In finalAttempted
                        If Not boundIds.Contains(kv.Key) Then
                            warn.Add("CAT_DROPPED_BY_REVIT: " & DescribeCategoryRef(kv.Value) & " (ResolvedId=" & kv.Key & ")")
                        End If
                    Next

                    If usedBindableOnly Then
                        For Each kv As KeyValuePair(Of Integer, CategoryRef) In attemptedFull
                            If Not finalAttempted.ContainsKey(kv.Key) Then
                                warn.Add("CAT_SKIPPED_BY_FALLBACK: " & DescribeCategoryRef(kv.Value) & " (ResolvedId=" & kv.Key & ")")
                            End If
                        Next
                    End If
                End If

                If warn.Count > 0 Then
                    warnLog = String.Join(vbCrLf, warn.ToArray())
                Else
                    warnLog = ""
                End If

                Return True

            Catch ex As Exception
                errLog = "트랜잭션 예외: " & ex.Message
                Return False
            End Try
        End Function

        Private Shared Function BuildBinding(app As RevitApp, cats As CategorySet, isInstance As Boolean) As Binding
            If isInstance Then
                Return app.Create.NewInstanceBinding(cats)
            Else
                Return app.Create.NewTypeBinding(cats)
            End If
        End Function

        Private Shared Function TryInsertOrReinsert(map As BindingMap,
                                                   extDef As ExternalDefinition,
                                                   binding As Binding,
                                                   pg As BuiltInParameterGroup) As Boolean
            If map Is Nothing OrElse extDef Is Nothing OrElse binding Is Nothing Then Return False

            Dim ok As Boolean = False
            Try
                ok = map.Insert(extDef, binding, pg)
                If Not ok Then
                    ok = map.ReInsert(extDef, binding, pg)
                End If
            Catch
                ok = False
            End Try
            Return ok
        End Function

        Private Shared Function DescribeCategoryRef(cref As CategoryRef) As String
            If cref Is Nothing Then Return "(null)"
            Dim pth As String = If(cref.Path, "").Trim()
            If pth <> "" Then
                Return pth & " (SavedId=" & cref.IdInt & ")"
            End If
            Dim n As String = If(cref.Name, "").Trim()
            If n <> "" Then
                Return n & " (SavedId=" & cref.IdInt & ")"
            End If
            Return "(SavedId=" & cref.IdInt & ")"
        End Function

        Private Shared Function ResolveCategoryInDoc(maps As CategoryMaps, cref As CategoryRef, ByRef resolvedBy As String) As Category
            resolvedBy = ""

            If maps Is Nothing OrElse cref Is Nothing Then
                resolvedBy = "none"
                Return Nothing
            End If

            Dim cat As Category = Nothing

            ' Prefer Path
            Dim pathStr As String = If(cref.Path, "").Trim()
            If pathStr <> "" Then
                If maps.ByPath.TryGetValue(pathStr, cat) AndAlso cat IsNot Nothing Then
                    resolvedBy = "path"
                    Return cat
                End If
            End If

            ' Fallback: Id
            If maps.ById.TryGetValue(cref.IdInt, cat) AndAlso cat IsNot Nothing Then
                resolvedBy = "id"
                Return cat
            End If

            resolvedBy = "notfound"
            Return Nothing
        End Function

        Private Shared Function GetBoundCategoryIds(doc As Document, extDefGuid As Guid) As HashSet(Of Integer)
            Dim hs As New HashSet(Of Integer)()

            Try
                Dim map As BindingMap = doc.ParameterBindings
                Dim it As DefinitionBindingMapIterator = map.ForwardIterator()
                it.Reset()

                While it.MoveNext()
                    Dim kDef As Definition = TryCast(it.Key, Definition)
                    Dim kExt As ExternalDefinition = TryCast(kDef, ExternalDefinition)

                    If kExt IsNot Nothing AndAlso kExt.GUID = extDefGuid Then
                        Dim eb As ElementBinding = TryCast(it.Current, ElementBinding)
                        If eb IsNot Nothing Then
                            Dim cs As CategorySet = eb.Categories
                            If cs IsNot Nothing Then
                                For Each c As Category In cs
                                    If c IsNot Nothing Then
                                        hs.Add(c.Id.IntegerValue)
                                    End If
                                Next
                            End If
                        End If
                        Exit While
                    End If
                End While
            Catch
            End Try

            Return hs
        End Function

        Private Shared Function HasParameterNameConflict(doc As Document, extDef As ExternalDefinition) As Boolean
            Dim spe As SharedParameterElement = SharedParameterElement.Lookup(doc, extDef.GUID)
            If spe IsNot Nothing Then
                Return False
            End If

            Try
                Dim collector As New FilteredElementCollector(doc)
                collector.OfClass(GetType(ParameterElement))
                For Each pe As ParameterElement In collector
                    If pe Is Nothing Then Continue For

                    Dim def As Definition = Nothing
                    Try
                        def = pe.GetDefinition()
                    Catch
                        def = Nothing
                    End Try
                    If def Is Nothing Then Continue For

                    If String.Equals(def.Name, extDef.Name, StringComparison.OrdinalIgnoreCase) Then
                        Dim speExisting As SharedParameterElement = TryCast(pe, SharedParameterElement)
                        If speExisting IsNot Nothing Then
                            If speExisting.GuidValue <> extDef.GUID Then
                                Return True
                            End If
                        Else
                            Return True
                        End If
                    End If
                Next
            Catch
                Return False
            End Try

            Return False
        End Function

        ' Vary Between Groups helper (apply after binding commit)
        Private Shared Function TrySetAllowVaryBetweenGroups(doc As Document, guid As Guid, allowVary As Boolean, ByRef warn As String) As Boolean
            warn = ""
            If doc Is Nothing Then
                warn = "VARY_SET_FAIL: doc is null"
                Return False
            End If

            Try
                Using t As New Autodesk.Revit.DB.Transaction(doc, "Set Vary Between Groups")
                    t.Start()

                    Dim spe As SharedParameterElement = SharedParameterElement.Lookup(doc, guid)
                    If spe Is Nothing Then
                        spe = FindSharedParameterElementByGuid(doc, guid)
                    End If

                    If spe Is Nothing Then
                        t.RollBack()
                        warn = "VARY_SET_FAIL: SharedParameterElement not found (GUID=" & guid.ToString() & ")"
                        Return False
                    End If

                    Dim idef As InternalDefinition = TryCast(spe.GetDefinition(), InternalDefinition)
                    If idef Is Nothing Then
                        t.RollBack()
                        warn = "VARY_SET_FAIL: InternalDefinition not found (GUID=" & guid.ToString() & ")"
                        Return False
                    End If

                    idef.SetAllowVaryBetweenGroups(doc, allowVary)

                    Try
                        doc.Regenerate()
                    Catch
                    End Try

                    t.Commit()
                End Using

                Return True

            Catch ex As Exception
                warn = "VARY_SET_FAIL: " & ex.Message
                Return False
            End Try
        End Function

        Private Shared Function FindSharedParameterElementByGuid(doc As Document, guid As Guid) As SharedParameterElement
            If doc Is Nothing Then Return Nothing
            Try
                Dim col As New FilteredElementCollector(doc)
                col.OfClass(GetType(SharedParameterElement))
                For Each el As Element In col
                    Dim spe As SharedParameterElement = TryCast(el, SharedParameterElement)
                    If spe Is Nothing Then Continue For
                    If spe.GuidValue = guid Then Return spe
                Next
            Catch
            End Try
            Return Nothing
        End Function

#End Region

#Region "Open / Sync / Close"

        Private Shared Function OpenProjectDocument(app As RevitApp,
                                                   userVisiblePath As String,
                                                   closeAllWorksets As Boolean) As Document
            Dim mp As ModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(userVisiblePath)

            Dim openOpts As New OpenOptions()
            openOpts.DetachFromCentralOption = DetachFromCentralOption.DoNotDetach

            If closeAllWorksets Then
                Dim wsc As New WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
                openOpts.SetOpenWorksetsConfiguration(wsc)
            End If

            Return app.OpenDocumentFile(mp, openOpts)
        End Function

        Private Shared Function SyncWithCentral(doc As Document,
                                               comment As String,
                                               ByRef err As String) As Boolean
            err = ""
            If doc Is Nothing OrElse Not doc.IsWorkshared Then
                err = "Workshared 문서가 아님."
                Return False
            End If

            Try
                Dim twc As New TransactWithCentralOptions()

                Dim swc As New SynchronizeWithCentralOptions()
                swc.Comment = If(comment, "")

                Try
                    Dim rel As New RelinquishOptions(True)
                    swc.SetRelinquishOptions(rel)
                Catch
                End Try

                doc.SynchronizeWithCentral(twc, swc)
                Return True

            Catch ex As Exception
                err = ex.Message
                Return False
            End Try
        End Function

        Private Shared Sub SafeClose(doc As Document, saveModified As Boolean)
            If doc Is Nothing Then Return
            Try
                doc.Close(saveModified)
            Catch
            End Try
        End Sub

        Private Shared Function IsAlreadyOpen(app As RevitApp, userVisiblePath As String) As Boolean
            Try
                For Each d As Document In app.Documents
                    If d Is Nothing Then Continue For
                    If String.Equals(d.PathName, userVisiblePath, StringComparison.OrdinalIgnoreCase) Then
                        Return True
                    End If
                Next
            Catch
                Return False
            End Try
            Return False
        End Function

#End Region

#Region "Categories Tree Build (include non-bindable leaves + Parent chain path)"

        ' Try-get Category even if Settings.Categories enumeration misses it in current doc state
        Private Shared Function TryGetCategoryByBuiltIn(doc As Document, bic As BuiltInCategory) As Category
            If doc Is Nothing Then Return Nothing

            Dim cat As Category = Nothing

            ' 1) Category.GetCategory 우선
            Try
                cat = Category.GetCategory(doc, bic)
                If cat IsNot Nothing Then Return cat
            Catch
                cat = Nothing
            End Try

            ' 2) Settings.Categories.Item(bic) 백업 (VB는 get_Item 아님)
            Try
                cat = doc.Settings.Categories.Item(bic)
                If cat IsNot Nothing Then Return cat
            Catch
                cat = Nothing
            End Try

            Return Nothing
        End Function

        ' Known missing candidates (keep minimal + targeted)
        Private Shared Sub EnsureExtraBuiltInCategories(doc As Document, dict As Dictionary(Of Integer, Category))
            If doc Is Nothing OrElse dict Is Nothing Then Return

            Dim bics As BuiltInCategory() = New BuiltInCategory() {
                BuiltInCategory.OST_Stairs,
                BuiltInCategory.OST_StairsRuns,
                BuiltInCategory.OST_StairsLandings,
                BuiltInCategory.OST_StairsSupports
            }

            For Each bic As BuiltInCategory In bics
                Dim c As Category = Nothing
                Try
                    c = TryGetCategoryByBuiltIn(doc, bic)
                Catch
                    c = Nothing
                End Try

                If c Is Nothing Then Continue For

                Dim idInt As Integer = c.Id.IntegerValue
                If Not dict.ContainsKey(idInt) Then
                    dict.Add(idInt, c)
                End If
            Next
        End Sub

        ' Build parent map primarily using SubCategories (more reliable than Category.Parent for some categories)
        Private Shared Function BuildParentMap(doc As Document, catsById As Dictionary(Of Integer, Category)) As Dictionary(Of Integer, Integer)
            Dim parentByChild As New Dictionary(Of Integer, Integer)()
            If catsById Is Nothing OrElse catsById.Count = 0 Then Return parentByChild

            ' 1) Primary: SubCategories graph
            For Each kv As KeyValuePair(Of Integer, Category) In catsById
                Dim parentCat As Category = kv.Value
                Dim pid As Integer = kv.Key
                If parentCat Is Nothing Then Continue For

                Try
                    Dim subs As CategoryNameMap = parentCat.SubCategories
                    If subs Is Nothing Then Continue For

                    For Each sc As Category In subs
                        If sc Is Nothing Then Continue For
                        Dim cid As Integer = sc.Id.IntegerValue
                        If cid = pid Then Continue For
                        If Not parentByChild.ContainsKey(cid) Then
                            parentByChild.Add(cid, pid)
                        End If
                    Next
                Catch
                End Try
            Next

            ' 2) Fallback: Category.Parent for unlinked items
            For Each kv As KeyValuePair(Of Integer, Category) In catsById
                Dim childCat As Category = kv.Value
                Dim cid As Integer = kv.Key
                If childCat Is Nothing Then Continue For
                If parentByChild.ContainsKey(cid) Then Continue For

                Try
                    Dim par As Category = childCat.Parent
                    If par Is Nothing Then Continue For
                    Dim pid As Integer = par.Id.IntegerValue
                    If pid <> cid AndAlso catsById.ContainsKey(pid) Then
                        parentByChild(cid) = pid
                    End If
                Catch
                End Try
            Next

            ' 3) Explicit fix: Stairs\Supports link (OST_StairsSupports -> OST_Stairs)
            Try
                Dim cSupports As Category = TryGetCategoryByBuiltIn(doc, BuiltInCategory.OST_StairsSupports)
                Dim cStairs As Category = TryGetCategoryByBuiltIn(doc, BuiltInCategory.OST_Stairs)
                If cSupports IsNot Nothing AndAlso cStairs IsNot Nothing Then
                    Dim cid As Integer = cSupports.Id.IntegerValue
                    Dim pid As Integer = cStairs.Id.IntegerValue
                    If cid <> pid Then
                        parentByChild(cid) = pid
                    End If
                End If
            Catch
            End Try

            Return parentByChild
        End Function

        Private Shared Function BuildCategoryTree(doc As Document) As List(Of CategoryTreeItem)
            Dim catsById As Dictionary(Of Integer, Category) = BuildFlatCategories(doc)
            If catsById.Count = 0 Then Return New List(Of CategoryTreeItem)()

            Dim parentMap As Dictionary(Of Integer, Integer) = BuildParentMap(doc, catsById)

            ' memoized path by id (Parent chain via parentMap)
            Dim pathCache As New Dictionary(Of Integer, String)()

            Dim nodesById As New Dictionary(Of Integer, CategoryTreeItem)()
            For Each kv As KeyValuePair(Of Integer, Category) In catsById
                Dim cat As Category = kv.Value
                If cat Is Nothing Then Continue For

                Dim bindable As Boolean = False
                Try
                    bindable = cat.AllowsBoundParameters
                Catch
                    bindable = False
                End Try

                ' ===== FIX: CType 키워드 충돌 제거 =====
                Dim catTypeVal As CategoryType = CategoryType.Model
                Try
                    catTypeVal = cat.CategoryType
                Catch
                    catTypeVal = CategoryType.Model
                End Try
                ' =======================================

                Dim node As New CategoryTreeItem() With {
                    .IdInt = kv.Key,
                    .Name = cat.Name,
                    .CatType = catTypeVal,
                    .IsBindable = bindable,
                    .Path = GetCategoryPath(kv.Key, catsById, parentMap, pathCache)
                }
                nodesById(kv.Key) = node
            Next

            ' link by parentMap (SubCategories-first)
            Dim roots As New List(Of CategoryTreeItem)()
            For Each kv As KeyValuePair(Of Integer, CategoryTreeItem) In nodesById
                Dim cid As Integer = kv.Key
                Dim node As CategoryTreeItem = kv.Value
                If node Is Nothing Then Continue For

                Dim pid As Integer = Integer.MinValue
                If parentMap IsNot Nothing AndAlso parentMap.TryGetValue(cid, pid) AndAlso pid <> cid AndAlso nodesById.ContainsKey(pid) Then
                    nodesById(pid).Children.Add(node)
                Else
                    roots.Add(node)
                End If
            Next

            ' remove duplicates in roots (safety)
            Dim uniq As New Dictionary(Of Integer, CategoryTreeItem)()
            For Each r As CategoryTreeItem In roots
                If r Is Nothing Then Continue For
                If Not uniq.ContainsKey(r.IdInt) Then uniq.Add(r.IdInt, r)
            Next

            Dim rootList As New List(Of CategoryTreeItem)()
            For Each v As CategoryTreeItem In uniq.Values
                rootList.Add(v)
            Next

            ' sort recursively
            SortCategoryTreeRecursive(rootList)

            Return rootList
        End Function

        Private Shared Sub SortCategoryTreeRecursive(list As List(Of CategoryTreeItem))
            If list Is Nothing Then Return
            list.Sort(Function(a, b) String.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase))
            For Each it As CategoryTreeItem In list
                If it.Children IsNot Nothing AndAlso it.Children.Count > 0 Then
                    SortCategoryTreeRecursive(it.Children)
                End If
            Next
        End Sub

        Private Shared Function BuildFlatCategories(doc As Document) As Dictionary(Of Integer, Category)
            Dim dict As New Dictionary(Of Integer, Category)()
            Try
                For Each c As Category In doc.Settings.Categories
                    CollectCategoryRecursive(c, dict)
                Next
            Catch
            End Try

            ' IMPORTANT: Ensure Stairs\Supports (and related) exist even if enumeration misses it
            EnsureExtraBuiltInCategories(doc, dict)

            Return dict
        End Function

        Private Shared Sub CollectCategoryRecursive(cat As Category, dict As Dictionary(Of Integer, Category))
            If cat Is Nothing OrElse dict Is Nothing Then Return

            Dim idInt As Integer = cat.Id.IntegerValue
            If Not dict.ContainsKey(idInt) Then
                dict.Add(idInt, cat)
            End If

            Try
                Dim subs As CategoryNameMap = cat.SubCategories
                If subs IsNot Nothing Then
                    For Each sc As Category In subs
                        CollectCategoryRecursive(sc, dict)
                    Next
                End If
            Catch
            End Try
        End Sub

        Private Shared Function GetCategoryPath(idInt As Integer,
                                               catsById As Dictionary(Of Integer, Category),
                                               parentMap As Dictionary(Of Integer, Integer),
                                               cache As Dictionary(Of Integer, String)) As String
            If cache IsNot Nothing AndAlso cache.ContainsKey(idInt) Then
                Return cache(idInt)
            End If

            Dim visiting As New HashSet(Of Integer)()
            Dim pathStr As String = GetCategoryPathInternal(idInt, catsById, parentMap, cache, visiting)
            Return pathStr
        End Function

        Private Shared Function GetCategoryPathInternal(idInt As Integer,
                                                       catsById As Dictionary(Of Integer, Category),
                                                       parentMap As Dictionary(Of Integer, Integer),
                                                       cache As Dictionary(Of Integer, String),
                                                       visiting As HashSet(Of Integer)) As String
            If cache IsNot Nothing AndAlso cache.ContainsKey(idInt) Then
                Return cache(idInt)
            End If

            If visiting.Contains(idInt) Then
                Dim nmLoop As String = ""
                If catsById IsNot Nothing AndAlso catsById.ContainsKey(idInt) AndAlso catsById(idInt) IsNot Nothing Then
                    nmLoop = catsById(idInt).Name
                Else
                    nmLoop = idInt.ToString()
                End If
                If cache IsNot Nothing Then cache(idInt) = nmLoop
                Return nmLoop
            End If

            visiting.Add(idInt)

            Dim cat As Category = Nothing
            If catsById Is Nothing OrElse Not catsById.TryGetValue(idInt, cat) OrElse cat Is Nothing Then
                visiting.Remove(idInt)
                If cache IsNot Nothing Then cache(idInt) = idInt.ToString()
                Return idInt.ToString()
            End If

            Dim result As String = cat.Name

            Dim pid As Integer = Integer.MinValue
            If parentMap IsNot Nothing AndAlso parentMap.TryGetValue(idInt, pid) Then
                If pid <> idInt AndAlso catsById.ContainsKey(pid) Then
                    Dim pPath As String = GetCategoryPathInternal(pid, catsById, parentMap, cache, visiting)
                    result = pPath & "\" & cat.Name
                End If
            End If

            visiting.Remove(idInt)

            If cache IsNot Nothing Then
                cache(idInt) = result
            End If

            Return result
        End Function

        Private Shared Function BuildAvailableCategoryMaps(doc As Document) As CategoryMaps
            Dim maps As New CategoryMaps()

            Dim catsById As Dictionary(Of Integer, Category) = BuildFlatCategories(doc)
            Dim parentMap As Dictionary(Of Integer, Integer) = BuildParentMap(doc, catsById)
            Dim pathCache As New Dictionary(Of Integer, String)()

            For Each kv As KeyValuePair(Of Integer, Category) In catsById
                Dim idInt As Integer = kv.Key
                Dim cat As Category = kv.Value
                If cat Is Nothing Then Continue For

                If Not maps.ById.ContainsKey(idInt) Then
                    maps.ById.Add(idInt, cat)
                End If

                Dim pth As String = GetCategoryPath(idInt, catsById, parentMap, pathCache)
                If Not String.IsNullOrWhiteSpace(pth) Then
                    If Not maps.ByPath.ContainsKey(pth) Then
                        maps.ByPath.Add(pth, cat)
                    End If
                End If
            Next

            Return maps
        End Function

#End Region

#Region "UI - Main Form"

        Friend Class SharedParamBatchMainForm
            Inherits WinForms.Form

            Private ReadOnly _app As RevitApp
            Private ReadOnly _spFilePath As String
            Private ReadOnly _categoryTree As List(Of CategoryTreeItem)

            Private _defFile As DefinitionFile
            Private _groupDefs As New Dictionary(Of String, List(Of ExternalDefinition))(StringComparer.OrdinalIgnoreCase)

            Private txtSpFile As WinForms.TextBox
            Private cmbGroup As WinForms.ComboBox
            Private lstDefs As WinForms.ListBox
            Private btnAddDefs As WinForms.Button

            Private dgvParams As WinForms.DataGridView
            Private btnApplyFirstToAll As WinForms.Button

            Private lstRvt As WinForms.ListBox
            Private btnAddRvt As WinForms.Button
            Private btnRemoveRvt As WinForms.Button
            Private btnClearRvt As WinForms.Button

            Private chkCloseAllWorksets As WinForms.CheckBox
            Private txtSyncComment As WinForms.TextBox

            Private btnOk As WinForms.Button
            Private btnCancel As WinForms.Button

            Private _paramList As New BindingList(Of ParamToBind)()
            Private _paramSource As New WinForms.BindingSource()

            Public Sub New(app As RevitApp, spFilePath As String, categoryTree As List(Of CategoryTreeItem))
                _app = app
                _spFilePath = spFilePath
                _categoryTree = categoryTree

                Me.Text = "Shared Parameter Batch Binder (Prototype)"
                Me.StartPosition = WinForms.FormStartPosition.CenterScreen
                Me.Width = 1200
                Me.Height = 820
                Me.MinimizeBox = False
                Me.MaximizeBox = False

                BuildUi()
                LoadSharedParameterFile()
                BindGrid()
            End Sub

            Public Function GetSettings() As BatchSettings
                Dim s As New BatchSettings()

                s.RvtPaths = New List(Of String)()
                For Each it As Object In lstRvt.Items
                    s.RvtPaths.Add(Convert.ToString(it))
                Next

                s.Parameters = _paramList.ToList()
                s.CloseAllWorksetsOnOpen = chkCloseAllWorksets.Checked
                s.SyncComment = txtSyncComment.Text

                Return s
            End Function

            Private Sub BuildUi()
                Dim main As New WinForms.TableLayoutPanel()
                main.Dock = WinForms.DockStyle.Fill
                main.ColumnCount = 2
                main.RowCount = 1
                main.ColumnStyles.Add(New WinForms.ColumnStyle(WinForms.SizeType.Percent, 55.0F))
                main.ColumnStyles.Add(New WinForms.ColumnStyle(WinForms.SizeType.Percent, 45.0F))
                Me.Controls.Add(main)

                Dim left As New WinForms.TableLayoutPanel()
                left.Dock = WinForms.DockStyle.Fill
                left.RowCount = 3
                left.ColumnCount = 1
                left.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Absolute, 80))
                left.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Absolute, 210))
                left.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Percent, 100))
                main.Controls.Add(left, 0, 0)

                Dim right As New WinForms.TableLayoutPanel()
                right.Dock = WinForms.DockStyle.Fill
                right.RowCount = 3
                right.ColumnCount = 1
                right.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Percent, 100))
                right.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Absolute, 120))
                right.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Absolute, 70))
                main.Controls.Add(right, 1, 0)

                Dim grpSpFile As New WinForms.GroupBox() With {.Text = "Shared Parameter Source (current Revit setting)", .Dock = WinForms.DockStyle.Fill}
                left.Controls.Add(grpSpFile, 0, 0)

                Dim spFileLayout As New WinForms.TableLayoutPanel()
                spFileLayout.Dock = WinForms.DockStyle.Fill
                spFileLayout.ColumnCount = 2
                spFileLayout.ColumnStyles.Add(New WinForms.ColumnStyle(WinForms.SizeType.Absolute, 110))
                spFileLayout.ColumnStyles.Add(New WinForms.ColumnStyle(WinForms.SizeType.Percent, 100))
                grpSpFile.Controls.Add(spFileLayout)

                spFileLayout.Controls.Add(New WinForms.Label() With {.Text = "TXT Path:", .AutoSize = True, .TextAlign = ContentAlignment.MiddleLeft}, 0, 0)
                txtSpFile = New WinForms.TextBox() With {.Dock = WinForms.DockStyle.Fill, .ReadOnly = True}
                txtSpFile.Text = _spFilePath
                spFileLayout.Controls.Add(txtSpFile, 1, 0)

                Dim hint As New WinForms.Label() With {
                    .Text = "※ 이 파일은 Revit(Manage > Shared Parameters)에서 설정된 값 그대로 사용합니다.",
                    .AutoSize = True
                }
                spFileLayout.Controls.Add(hint, 1, 1)

                Dim grpAvail As New WinForms.GroupBox() With {.Text = "Select Shared Parameters to add (1+)", .Dock = WinForms.DockStyle.Fill}
                left.Controls.Add(grpAvail, 0, 1)

                Dim avail As New WinForms.TableLayoutPanel()
                avail.Dock = WinForms.DockStyle.Fill
                avail.ColumnCount = 3
                avail.RowCount = 3
                avail.ColumnStyles.Add(New WinForms.ColumnStyle(WinForms.SizeType.Absolute, 90))
                avail.ColumnStyles.Add(New WinForms.ColumnStyle(WinForms.SizeType.Percent, 100))
                avail.ColumnStyles.Add(New WinForms.ColumnStyle(WinForms.SizeType.Absolute, 140))
                avail.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Absolute, 30))
                avail.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Percent, 100))
                avail.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Absolute, 40))
                grpAvail.Controls.Add(avail)

                avail.Controls.Add(New WinForms.Label() With {.Text = "Group:", .AutoSize = True}, 0, 0)
                cmbGroup = New WinForms.ComboBox() With {.Dock = WinForms.DockStyle.Fill, .DropDownStyle = WinForms.ComboBoxStyle.DropDownList}
                AddHandler cmbGroup.SelectedIndexChanged, AddressOf OnGroupChanged
                avail.Controls.Add(cmbGroup, 1, 0)

                btnAddDefs = New WinForms.Button() With {.Text = "Add Selected →", .Width = 130, .Height = 28}
                AddHandler btnAddDefs.Click, AddressOf OnAddSelectedDefs
                avail.Controls.Add(btnAddDefs, 2, 0)

                lstDefs = New WinForms.ListBox() With {.Dock = WinForms.DockStyle.Fill, .SelectionMode = WinForms.SelectionMode.MultiExtended}
                avail.Controls.Add(lstDefs, 0, 1)
                avail.SetColumnSpan(lstDefs, 3)

                Dim lblTip As New WinForms.Label() With {.Text = "Tip: Ctrl/Shift로 다중 선택 후 Add", .AutoSize = True}
                avail.Controls.Add(lblTip, 0, 2)
                avail.SetColumnSpan(lblTip, 2)

                Dim grpSel As New WinForms.GroupBox() With {.Text = "Selected Parameters (each has its own settings)", .Dock = WinForms.DockStyle.Fill}
                left.Controls.Add(grpSel, 0, 2)

                Dim sel As New WinForms.TableLayoutPanel()
                sel.Dock = WinForms.DockStyle.Fill
                sel.RowCount = 2
                sel.ColumnCount = 1
                sel.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Absolute, 40))
                sel.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Percent, 100))
                grpSel.Controls.Add(sel)

                Dim selTop As New WinForms.FlowLayoutPanel() With {.Dock = WinForms.DockStyle.Fill, .FlowDirection = WinForms.FlowDirection.LeftToRight}
                btnApplyFirstToAll = New WinForms.Button() With {.Text = "Apply FIRST parameter settings to ALL", .Width = 260, .Height = 28}
                AddHandler btnApplyFirstToAll.Click, AddressOf OnApplyFirstToAll
                selTop.Controls.Add(btnApplyFirstToAll)
                sel.Controls.Add(selTop, 0, 0)

                dgvParams = New WinForms.DataGridView()
                dgvParams.Dock = WinForms.DockStyle.Fill
                dgvParams.AllowUserToAddRows = False
                dgvParams.AllowUserToDeleteRows = False
                dgvParams.ReadOnly = True
                dgvParams.AutoGenerateColumns = False
                dgvParams.SelectionMode = WinForms.DataGridViewSelectionMode.FullRowSelect
                dgvParams.MultiSelect = False
                dgvParams.RowHeadersVisible = False
                AddHandler dgvParams.CellContentClick, AddressOf OnParamGridCellClick
                sel.Controls.Add(dgvParams, 0, 1)

                Dim grpRvt As New WinForms.GroupBox() With {.Text = "Target RVT files (batch)", .Dock = WinForms.DockStyle.Fill}
                right.Controls.Add(grpRvt, 0, 0)

                Dim rvt As New WinForms.TableLayoutPanel()
                rvt.Dock = WinForms.DockStyle.Fill
                rvt.ColumnCount = 1
                rvt.RowCount = 2
                rvt.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Percent, 100))
                rvt.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Absolute, 40))
                grpRvt.Controls.Add(rvt)

                lstRvt = New WinForms.ListBox() With {.Dock = WinForms.DockStyle.Fill}
                rvt.Controls.Add(lstRvt, 0, 0)

                Dim rvtBtns As New WinForms.FlowLayoutPanel() With {.Dock = WinForms.DockStyle.Fill, .FlowDirection = WinForms.FlowDirection.LeftToRight}
                btnAddRvt = New WinForms.Button() With {.Text = "Add RVT(s)...", .Width = 120}
                btnRemoveRvt = New WinForms.Button() With {.Text = "Remove", .Width = 90}
                btnClearRvt = New WinForms.Button() With {.Text = "Clear", .Width = 90}
                AddHandler btnAddRvt.Click, AddressOf OnAddRvts
                AddHandler btnRemoveRvt.Click, AddressOf OnRemoveRvt
                AddHandler btnClearRvt.Click, AddressOf OnClearRvt
                rvtBtns.Controls.Add(btnAddRvt)
                rvtBtns.Controls.Add(btnRemoveRvt)
                rvtBtns.Controls.Add(btnClearRvt)
                rvt.Controls.Add(rvtBtns, 0, 1)

                Dim grpOpt As New WinForms.GroupBox() With {.Text = "Options", .Dock = WinForms.DockStyle.Fill}
                right.Controls.Add(grpOpt, 0, 1)

                Dim opt As New WinForms.TableLayoutPanel()
                opt.Dock = WinForms.DockStyle.Fill
                opt.ColumnCount = 2
                opt.RowCount = 3
                opt.ColumnStyles.Add(New WinForms.ColumnStyle(WinForms.SizeType.Absolute, 140))
                opt.ColumnStyles.Add(New WinForms.ColumnStyle(WinForms.SizeType.Percent, 100))
                opt.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Absolute, 30))
                opt.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Absolute, 30))
                opt.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Absolute, 30))
                grpOpt.Controls.Add(opt)

                chkCloseAllWorksets = New WinForms.CheckBox() With {.Text = "Workshared: Open CloseAllWorksets", .AutoSize = True, .Checked = True}
                opt.Controls.Add(chkCloseAllWorksets, 0, 0)
                opt.SetColumnSpan(chkCloseAllWorksets, 2)

                opt.Controls.Add(New WinForms.Label() With {.Text = "Sync Comment:", .AutoSize = True}, 0, 1)
                txtSyncComment = New WinForms.TextBox() With {.Dock = WinForms.DockStyle.Fill}
                opt.Controls.Add(txtSyncComment, 1, 1)

                Dim info As New WinForms.Label() With {.Text = "※ Workshared 모델만 Sync Comment가 남습니다.", .AutoSize = True}
                opt.Controls.Add(info, 1, 2)

                Dim bottom As New WinForms.FlowLayoutPanel() With {.Dock = WinForms.DockStyle.Fill, .FlowDirection = WinForms.FlowDirection.RightToLeft}
                right.Controls.Add(bottom, 0, 2)

                btnOk = New WinForms.Button() With {.Text = "Run", .Width = 110, .Height = 32}
                btnCancel = New WinForms.Button() With {.Text = "Cancel", .Width = 110, .Height = 32}
                AddHandler btnOk.Click, AddressOf OnOk
                AddHandler btnCancel.Click, Sub(sender, e) Me.DialogResult = WinForms.DialogResult.Cancel
                bottom.Controls.Add(btnOk)
                bottom.Controls.Add(btnCancel)
            End Sub

            Private Sub LoadSharedParameterFile()
                _groupDefs.Clear()
                cmbGroup.Items.Clear()
                lstDefs.Items.Clear()

                Try
                    Dim prev As String = _app.SharedParametersFilename
                    _app.SharedParametersFilename = _spFilePath
                    _defFile = _app.OpenSharedParameterFile()
                    _app.SharedParametersFilename = prev

                    If _defFile Is Nothing Then
                        WinForms.MessageBox.Show("Shared Parameter 파일을 열 수 없습니다: " & _spFilePath, "Error",
                                                 WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error)
                        Return
                    End If

                    For Each g As DefinitionGroup In _defFile.Groups
                        If g Is Nothing Then Continue For
                        Dim list As New List(Of ExternalDefinition)()
                        For Each d As Definition In g.Definitions
                            Dim ed As ExternalDefinition = TryCast(d, ExternalDefinition)
                            If ed IsNot Nothing Then list.Add(ed)
                        Next
                        list = list.OrderBy(Function(x) x.Name).ToList()
                        _groupDefs(g.Name) = list
                    Next

                    For Each gName As String In _groupDefs.Keys.OrderBy(Function(x) x)
                        cmbGroup.Items.Add(gName)
                    Next
                    If cmbGroup.Items.Count > 0 Then cmbGroup.SelectedIndex = 0

                Catch ex As Exception
                    WinForms.MessageBox.Show("Shared Parameter 로드 예외: " & ex.Message, "Error",
                                             WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error)
                End Try
            End Sub

            Private Sub BindGrid()
                dgvParams.Columns.Clear()

                dgvParams.Columns.Add(New WinForms.DataGridViewTextBoxColumn() With {.HeaderText = "Name", .DataPropertyName = "ParamName", .Width = 150})
                dgvParams.Columns.Add(New WinForms.DataGridViewTextBoxColumn() With {.HeaderText = "Group", .DataPropertyName = "GroupName", .Width = 140})
                dgvParams.Columns.Add(New WinForms.DataGridViewTextBoxColumn() With {.HeaderText = "Type", .DataPropertyName = "ParamTypeLabel", .Width = 120})
                dgvParams.Columns.Add(New WinForms.DataGridViewTextBoxColumn() With {.HeaderText = "GUID", .DataPropertyName = "GuidString", .Width = 240})
                dgvParams.Columns.Add(New WinForms.DataGridViewTextBoxColumn() With {.HeaderText = "Binding", .DataPropertyName = "BindingDisplay", .Width = 70})
                dgvParams.Columns.Add(New WinForms.DataGridViewTextBoxColumn() With {.HeaderText = "Group Under", .DataPropertyName = "ParamGroupDisplay", .Width = 160})
                dgvParams.Columns.Add(New WinForms.DataGridViewTextBoxColumn() With {.HeaderText = "Cats", .DataPropertyName = "CategoriesCountDisplay", .Width = 50})

                Dim btnSettings As New WinForms.DataGridViewButtonColumn()
                btnSettings.HeaderText = ""
                btnSettings.Text = "Settings..."
                btnSettings.UseColumnTextForButtonValue = True
                btnSettings.Width = 90
                dgvParams.Columns.Add(btnSettings)

                Dim btnRemove As New WinForms.DataGridViewButtonColumn()
                btnRemove.HeaderText = ""
                btnRemove.Text = "Remove"
                btnRemove.UseColumnTextForButtonValue = True
                btnRemove.Width = 80
                dgvParams.Columns.Add(btnRemove)

                _paramSource.DataSource = _paramList
                dgvParams.DataSource = _paramSource
            End Sub

            Private Sub RefreshGrid()
                Try
                    _paramSource.ResetBindings(False)
                Catch
                    dgvParams.Refresh()
                End Try
            End Sub

            Private Sub OnGroupChanged(sender As Object, e As EventArgs)
                lstDefs.Items.Clear()
                Dim gName As String = Convert.ToString(cmbGroup.SelectedItem)
                If String.IsNullOrWhiteSpace(gName) Then Return
                If Not _groupDefs.ContainsKey(gName) Then Return

                For Each ed As ExternalDefinition In _groupDefs(gName)
                    lstDefs.Items.Add(New ExtDefItem(ed))
                Next
            End Sub

            Private Sub OnAddSelectedDefs(sender As Object, e As EventArgs)
                If lstDefs.SelectedItems Is Nothing OrElse lstDefs.SelectedItems.Count = 0 Then Return

                For Each obj As Object In lstDefs.SelectedItems
                    Dim it As ExtDefItem = TryCast(obj, ExtDefItem)
                    If it Is Nothing OrElse it.Def Is Nothing Then Continue For

                    Dim guid As Guid = it.Def.GUID
                    If _paramList.Any(Function(x) x.GuidValue = guid) Then
                        Continue For
                    End If

                    Dim typeLabel As String = ""
                    Try
                        typeLabel = LabelUtils.GetLabelFor(it.Def.ParameterType)
                    Catch
                        typeLabel = it.Def.ParameterType.ToString()
                    End Try

                    Dim p As New ParamToBind() With {
                        .GroupName = Convert.ToString(cmbGroup.SelectedItem),
                        .ParamName = it.Def.Name,
                        .GuidValue = guid,
                        .ParamTypeLabel = typeLabel,
                        .Description = If(it.Def.Description, ""),
                        .Settings = New ParamBindingSettings() With {
                            .IsInstanceBinding = True,
                            .ParamGroup = BuiltInParameterGroup.PG_DATA,
                            .AllowVaryBetweenGroups = False,
                            .Categories = New List(Of CategoryRef)()
                        }
                    }
                    _paramList.Add(p)
                Next

                RefreshGrid()
            End Sub

            Private Sub OnParamGridCellClick(sender As Object, e As WinForms.DataGridViewCellEventArgs)
                If e.RowIndex < 0 Then Return
                If e.ColumnIndex < 0 Then Return

                Dim row As WinForms.DataGridViewRow = dgvParams.Rows(e.RowIndex)
                Dim p As ParamToBind = TryCast(row.DataBoundItem, ParamToBind)
                If p Is Nothing Then Return

                Dim settingsColIndex As Integer = dgvParams.Columns.Count - 2
                Dim removeColIndex As Integer = dgvParams.Columns.Count - 1

                If e.ColumnIndex = settingsColIndex Then
                    Using dlg As New ParamSettingsForm(p, _categoryTree)
                        Dim dr As WinForms.DialogResult = dlg.ShowDialog()
                        If dr = WinForms.DialogResult.OK Then
                            p.Settings = dlg.GetSettings()
                            RefreshGrid()
                        End If
                    End Using

                ElseIf e.ColumnIndex = removeColIndex Then
                    _paramList.Remove(p)
                    RefreshGrid()
                End If
            End Sub

            Private Sub OnApplyFirstToAll(sender As Object, e As EventArgs)
                If _paramList.Count <= 1 Then Return
                Dim src As ParamToBind = _paramList(0)
                If src Is Nothing OrElse src.Settings Is Nothing Then Return

                Dim clone As ParamBindingSettings = src.Settings.CloneDeep()
                For i As Integer = 1 To _paramList.Count - 1
                    _paramList(i).Settings = clone.CloneDeep()
                Next

                RefreshGrid()
            End Sub

            Private Sub OnAddRvts(sender As Object, e As EventArgs)
                Using ofd As New WinForms.OpenFileDialog()
                    ofd.Filter = "Revit Project (*.rvt)|*.rvt"
                    ofd.Multiselect = True
                    ofd.Title = "Select RVT files"
                    If ofd.ShowDialog() = WinForms.DialogResult.OK Then
                        For Each p As String In ofd.FileNames
                            If Not lstRvt.Items.Contains(p) Then lstRvt.Items.Add(p)
                        Next
                    End If
                End Using
            End Sub

            Private Sub OnRemoveRvt(sender As Object, e As EventArgs)
                Dim idx As Integer = lstRvt.SelectedIndex
                If idx >= 0 Then lstRvt.Items.RemoveAt(idx)
            End Sub

            Private Sub OnClearRvt(sender As Object, e As EventArgs)
                lstRvt.Items.Clear()
            End Sub

            Private Sub OnOk(sender As Object, e As EventArgs)
                Dim s As BatchSettings = GetSettings()

                If s.RvtPaths Is Nothing OrElse s.RvtPaths.Count = 0 Then
                    WinForms.MessageBox.Show("대상 RVT 파일을 1개 이상 선택하세요.", "Validation",
                                             WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning)
                    Return
                End If
                If s.Parameters Is Nothing OrElse s.Parameters.Count = 0 Then
                    WinForms.MessageBox.Show("추가할 Shared Parameter를 1개 이상 등록하세요.", "Validation",
                                             WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning)
                    Return
                End If

                Dim noCat As New List(Of String)()
                For Each p As ParamToBind In s.Parameters
                    If p Is Nothing OrElse p.Settings Is Nothing OrElse p.Settings.Categories Is Nothing OrElse p.Settings.Categories.Count = 0 Then
                        noCat.Add(If(p IsNot Nothing, p.ParamName, "(unknown)"))
                    End If
                Next
                If noCat.Count > 0 Then
                    WinForms.MessageBox.Show("카테고리를 지정하지 않은 파라미터가 있습니다:" & vbCrLf &
                                             String.Join(", ", noCat.ToArray()),
                                             "Validation", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning)
                    Return
                End If

                Me.DialogResult = WinForms.DialogResult.OK
            End Sub

            Private Class ExtDefItem
                Public ReadOnly Property Def As ExternalDefinition
                Public Sub New(d As ExternalDefinition)
                    Def = d
                End Sub
                Public Overrides Function ToString() As String
                    If Def Is Nothing Then Return "(null)"
                    Return Def.Name
                End Function
            End Class

        End Class

#End Region

#Region "UI - Parameter Settings Form"

        Friend Class ParamSettingsForm
            Inherits WinForms.Form

            Private ReadOnly _param As ParamToBind
            Private ReadOnly _categoryTree As List(Of CategoryTreeItem)

            Private ReadOnly _checked As New Dictionary(Of Integer, Boolean)()
            Private ReadOnly _idToPath As New Dictionary(Of Integer, String)()
            Private ReadOnly _idToName As New Dictionary(Of Integer, String)()

            Private _suppressTreeEvents As Boolean = False

            Private lblName As WinForms.Label
            Private lblType As WinForms.Label
            Private lblGuid As WinForms.Label
            Private txtDesc As WinForms.TextBox

            Private rdoInstance As WinForms.RadioButton
            Private rdoType As WinForms.RadioButton

            Private cmbParamGroup As WinForms.ComboBox

            Private rdoVaryAligned As WinForms.RadioButton
            Private rdoVaryEach As WinForms.RadioButton

            Private cmbFilter As WinForms.ComboBox
            Private chkHideUnchecked As WinForms.CheckBox
            Private tvCats As WinForms.TreeView
            Private btnCheckAll As WinForms.Button
            Private btnCheckNone As WinForms.Button

            Private btnOk As WinForms.Button
            Private btnCancel As WinForms.Button

            Public Sub New(p As ParamToBind, categoryTree As List(Of CategoryTreeItem))
                _param = p
                _categoryTree = categoryTree

                Me.Text = "Parameter Settings - " & If(p IsNot Nothing, p.ParamName, "")
                Me.StartPosition = WinForms.FormStartPosition.CenterParent
                Me.Width = 980
                Me.Height = 720
                Me.MinimizeBox = False
                Me.MaximizeBox = False

                BuildUi()
                FillParamGroupCombo()

                BuildIdMaps()
                InitFromParam()
                RebuildTree()
            End Sub

            Public Function GetSettings() As ParamBindingSettings
                Dim s As New ParamBindingSettings()
                s.IsInstanceBinding = rdoInstance.Checked
                s.ParamGroup = GetSelectedParamGroup()
                s.AllowVaryBetweenGroups = (rdoInstance.Checked AndAlso rdoVaryEach.Checked)

                Dim list As New List(Of CategoryRef)()
                For Each kv As KeyValuePair(Of Integer, Boolean) In _checked
                    If Not kv.Value Then Continue For
                    Dim idInt As Integer = kv.Key
                    Dim pathStr As String = ""
                    Dim nameStr As String = ""
                    If _idToPath.ContainsKey(idInt) Then pathStr = _idToPath(idInt)
                    If _idToName.ContainsKey(idInt) Then nameStr = _idToName(idInt)
                    list.Add(New CategoryRef() With {.IdInt = idInt, .Path = pathStr, .Name = nameStr})
                Next
                s.Categories = list

                Return s
            End Function

            Private Sub BuildIdMaps()
                _idToPath.Clear()
                _idToName.Clear()
                For Each root As CategoryTreeItem In _categoryTree
                    CollectIdMaps(root)
                Next
            End Sub

            Private Sub CollectIdMaps(it As CategoryTreeItem)
                If it Is Nothing Then Return
                If Not _idToPath.ContainsKey(it.IdInt) Then _idToPath.Add(it.IdInt, it.Path)
                If Not _idToName.ContainsKey(it.IdInt) Then _idToName.Add(it.IdInt, it.Name)
                If it.Children Is Nothing Then Return
                For Each ch As CategoryTreeItem In it.Children
                    CollectIdMaps(ch)
                Next
            End Sub

            Private Sub BuildUi()
                Dim main As New WinForms.TableLayoutPanel()
                main.Dock = WinForms.DockStyle.Fill
                main.ColumnCount = 2
                main.RowCount = 2
                main.ColumnStyles.Add(New WinForms.ColumnStyle(WinForms.SizeType.Percent, 45.0F))
                main.ColumnStyles.Add(New WinForms.ColumnStyle(WinForms.SizeType.Percent, 55.0F))
                main.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Percent, 100.0F))
                main.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Absolute, 60.0F))
                Me.Controls.Add(main)

                Dim left As New WinForms.TableLayoutPanel()
                left.Dock = WinForms.DockStyle.Fill
                left.RowCount = 2
                left.ColumnCount = 1
                left.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Absolute, 220))
                left.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Percent, 100))
                main.Controls.Add(left, 0, 0)

                Dim grpInfo As New WinForms.GroupBox() With {.Text = "Shared Parameter Info", .Dock = WinForms.DockStyle.Fill}
                left.Controls.Add(grpInfo, 0, 0)

                Dim info As New WinForms.TableLayoutPanel()
                info.Dock = WinForms.DockStyle.Fill
                info.ColumnCount = 2
                info.RowCount = 4
                info.ColumnStyles.Add(New WinForms.ColumnStyle(WinForms.SizeType.Absolute, 90))
                info.ColumnStyles.Add(New WinForms.ColumnStyle(WinForms.SizeType.Percent, 100))
                info.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Absolute, 28))
                info.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Absolute, 28))
                info.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Absolute, 28))
                info.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Percent, 100))
                grpInfo.Controls.Add(info)

                info.Controls.Add(New WinForms.Label() With {.Text = "Name:", .AutoSize = True}, 0, 0)
                lblName = New WinForms.Label() With {.AutoSize = True}
                info.Controls.Add(lblName, 1, 0)

                info.Controls.Add(New WinForms.Label() With {.Text = "Type:", .AutoSize = True}, 0, 1)
                lblType = New WinForms.Label() With {.AutoSize = True}
                info.Controls.Add(lblType, 1, 1)

                info.Controls.Add(New WinForms.Label() With {.Text = "GUID:", .AutoSize = True}, 0, 2)
                lblGuid = New WinForms.Label() With {.AutoSize = True}
                info.Controls.Add(lblGuid, 1, 2)

                info.Controls.Add(New WinForms.Label() With {.Text = "Description:", .AutoSize = True}, 0, 3)
                txtDesc = New WinForms.TextBox() With {.Dock = WinForms.DockStyle.Fill, .Multiline = True, .ReadOnly = True, .ScrollBars = WinForms.ScrollBars.Vertical}
                info.Controls.Add(txtDesc, 1, 3)

                Dim grpBind As New WinForms.GroupBox() With {.Text = "Binding Options", .Dock = WinForms.DockStyle.Fill}
                left.Controls.Add(grpBind, 0, 1)

                Dim bind As New WinForms.TableLayoutPanel()
                bind.Dock = WinForms.DockStyle.Fill
                bind.ColumnCount = 2
                bind.RowCount = 5
                bind.ColumnStyles.Add(New WinForms.ColumnStyle(WinForms.SizeType.Absolute, 170))
                bind.ColumnStyles.Add(New WinForms.ColumnStyle(WinForms.SizeType.Percent, 100))
                bind.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Absolute, 34))
                bind.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Absolute, 34))
                bind.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Absolute, 60))
                bind.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Absolute, 34))
                bind.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Percent, 100))
                grpBind.Controls.Add(bind)

                bind.Controls.Add(New WinForms.Label() With {.Text = "Type/Instance:", .AutoSize = True}, 0, 0)
                Dim pnlTI As New WinForms.FlowLayoutPanel() With {.Dock = WinForms.DockStyle.Fill}
                rdoInstance = New WinForms.RadioButton() With {.Text = "Instance", .AutoSize = True, .Checked = True}
                rdoType = New WinForms.RadioButton() With {.Text = "Type", .AutoSize = True}
                AddHandler rdoInstance.CheckedChanged, AddressOf OnBindingTypeChanged
                AddHandler rdoType.CheckedChanged, AddressOf OnBindingTypeChanged
                pnlTI.Controls.Add(rdoInstance)
                pnlTI.Controls.Add(rdoType)
                bind.Controls.Add(pnlTI, 1, 0)

                bind.Controls.Add(New WinForms.Label() With {.Text = "Group parameter under:", .AutoSize = True}, 0, 1)
                cmbParamGroup = New WinForms.ComboBox() With {.Dock = WinForms.DockStyle.Fill, .DropDownStyle = WinForms.ComboBoxStyle.DropDownList}
                bind.Controls.Add(cmbParamGroup, 1, 1)

                bind.Controls.Add(New WinForms.Label() With {.Text = "Group values:", .AutoSize = True}, 0, 2)
                Dim pnlVary As New WinForms.FlowLayoutPanel() With {.Dock = WinForms.DockStyle.Fill, .FlowDirection = WinForms.FlowDirection.TopDown, .WrapContents = False}
                rdoVaryAligned = New WinForms.RadioButton() With {.Text = "Values are aligned per group type", .AutoSize = True, .Checked = True}
                rdoVaryEach = New WinForms.RadioButton() With {.Text = "Values can vary by group instance", .AutoSize = True}
                pnlVary.Controls.Add(rdoVaryAligned)
                pnlVary.Controls.Add(rdoVaryEach)
                bind.Controls.Add(pnlVary, 1, 2)

                Dim note As New WinForms.Label() With {.Text = "※ Instance일 때만 활성화됩니다.", .AutoSize = True}
                bind.Controls.Add(note, 1, 3)
                bind.SetColumnSpan(note, 2)

                Dim grpCats As New WinForms.GroupBox() With {.Text = "Categories (tree with sub-categories)", .Dock = WinForms.DockStyle.Fill}
                main.Controls.Add(grpCats, 1, 0)

                Dim cat As New WinForms.TableLayoutPanel()
                cat.Dock = WinForms.DockStyle.Fill
                cat.ColumnCount = 3
                cat.RowCount = 4
                cat.ColumnStyles.Add(New WinForms.ColumnStyle(WinForms.SizeType.Absolute, 100))
                cat.ColumnStyles.Add(New WinForms.ColumnStyle(WinForms.SizeType.Percent, 100))
                cat.ColumnStyles.Add(New WinForms.ColumnStyle(WinForms.SizeType.Absolute, 220))
                cat.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Absolute, 34))
                cat.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Absolute, 34))
                cat.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Percent, 100))
                cat.RowStyles.Add(New WinForms.RowStyle(WinForms.SizeType.Absolute, 44))
                grpCats.Controls.Add(cat)

                cat.Controls.Add(New WinForms.Label() With {.Text = "Filter list:", .AutoSize = True}, 0, 0)
                cmbFilter = New WinForms.ComboBox() With {.Dock = WinForms.DockStyle.Fill, .DropDownStyle = WinForms.ComboBoxStyle.DropDownList}
                cmbFilter.Items.Add("<show all>")
                cmbFilter.Items.Add("Model")
                cmbFilter.Items.Add("Annotation")
                cmbFilter.Items.Add("Analytical")
                cmbFilter.Items.Add("Internal")
                cmbFilter.SelectedIndex = 0
                AddHandler cmbFilter.SelectedIndexChanged, Sub(s, e) RebuildTree()
                cat.Controls.Add(cmbFilter, 1, 0)

                chkHideUnchecked = New WinForms.CheckBox() With {.Text = "Hide un-checked categories", .AutoSize = True}
                AddHandler chkHideUnchecked.CheckedChanged, Sub(s, e) RebuildTree()
                cat.Controls.Add(chkHideUnchecked, 2, 0)

                tvCats = New WinForms.TreeView() With {
                    .Dock = WinForms.DockStyle.Fill,
                    .CheckBoxes = True,
                    .ShowLines = True,
                    .ShowPlusMinus = True,
                    .ShowRootLines = True,
                    .ShowNodeToolTips = True
                }
                AddHandler tvCats.AfterCheck, AddressOf OnAfterCheckTree
                cat.Controls.Add(tvCats, 0, 1)
                cat.SetColumnSpan(tvCats, 3)
                cat.SetRowSpan(tvCats, 2)

                Dim pnlBtns As New WinForms.FlowLayoutPanel() With {.Dock = WinForms.DockStyle.Fill}
                btnCheckAll = New WinForms.Button() With {.Text = "Check All", .Width = 100}
                btnCheckNone = New WinForms.Button() With {.Text = "Check None", .Width = 100}
                AddHandler btnCheckAll.Click, Sub(s, e) SetAllChecked(True)
                AddHandler btnCheckNone.Click, Sub(s, e) SetAllChecked(False)
                pnlBtns.Controls.Add(btnCheckAll)
                pnlBtns.Controls.Add(btnCheckNone)
                cat.Controls.Add(pnlBtns, 2, 3)

                Dim bottom As New WinForms.FlowLayoutPanel() With {.Dock = WinForms.DockStyle.Fill, .FlowDirection = WinForms.FlowDirection.RightToLeft}
                main.Controls.Add(bottom, 0, 1)
                main.SetColumnSpan(bottom, 2)

                btnOk = New WinForms.Button() With {.Text = "OK", .Width = 110, .Height = 32}
                btnCancel = New WinForms.Button() With {.Text = "Cancel", .Width = 110, .Height = 32}
                AddHandler btnOk.Click, AddressOf OnOk
                AddHandler btnCancel.Click, Sub(sender, e) Me.DialogResult = WinForms.DialogResult.Cancel

                bottom.Controls.Add(btnOk)
                bottom.Controls.Add(btnCancel)
            End Sub

            Private Sub FillParamGroupCombo()
                cmbParamGroup.Items.Clear()

                Dim items As New List(Of ParamGroupItem)()

                Dim vals As Array = System.Enum.GetValues(GetType(BuiltInParameterGroup))
                For Each vObj As Object In vals
                    Dim v As BuiltInParameterGroup = CType(vObj, BuiltInParameterGroup)

                    Dim label As String = ""
                    Try
                        label = LabelUtils.GetLabelFor(v)
                    Catch
                        label = v.ToString()
                    End Try

                    items.Add(New ParamGroupItem(v, label))
                Next

                items.Sort(AddressOf CompareParamGroupItem)

                For Each it As ParamGroupItem In items
                    cmbParamGroup.Items.Add(it)
                Next

                Dim defaultItem As ParamGroupItem = Nothing
                For Each it As ParamGroupItem In items
                    If it.Value = BuiltInParameterGroup.PG_DATA Then
                        defaultItem = it
                        Exit For
                    End If
                Next

                If defaultItem IsNot Nothing Then
                    cmbParamGroup.SelectedItem = defaultItem
                ElseIf cmbParamGroup.Items.Count > 0 Then
                    cmbParamGroup.SelectedIndex = 0
                End If
            End Sub

            Private Shared Function CompareParamGroupItem(a As ParamGroupItem, b As ParamGroupItem) As Integer
                If a Is Nothing AndAlso b Is Nothing Then Return 0
                If a Is Nothing Then Return -1
                If b Is Nothing Then Return 1
                Return String.Compare(a.Label, b.Label, StringComparison.OrdinalIgnoreCase)
            End Function

            Private Function GetSelectedParamGroup() As BuiltInParameterGroup
                Dim it As ParamGroupItem = TryCast(cmbParamGroup.SelectedItem, ParamGroupItem)
                If it IsNot Nothing Then Return it.Value
                Return BuiltInParameterGroup.PG_DATA
            End Function

            Private Class ParamGroupItem
                Public ReadOnly Property Value As BuiltInParameterGroup
                Public ReadOnly Property Label As String
                Public Sub New(v As BuiltInParameterGroup, label As String)
                    Me.Value = v
                    Me.Label = label
                End Sub
                Public Overrides Function ToString() As String
                    Return Label
                End Function
            End Class

            Private Sub InitFromParam()
                If _param Is Nothing Then Return

                lblName.Text = _param.ParamName
                lblType.Text = _param.ParamTypeLabel
                lblGuid.Text = _param.GuidValue.ToString()
                txtDesc.Text = If(_param.Description, "")

                _checked.Clear()

                If _param.Settings IsNot Nothing AndAlso _param.Settings.Categories IsNot Nothing Then
                    For Each cr As CategoryRef In _param.Settings.Categories
                        If cr Is Nothing Then Continue For
                        _checked(cr.IdInt) = True
                    Next
                End If

                If _param.Settings IsNot Nothing Then
                    rdoInstance.Checked = _param.Settings.IsInstanceBinding
                    rdoType.Checked = Not _param.Settings.IsInstanceBinding

                    Dim target As BuiltInParameterGroup = _param.Settings.ParamGroup
                    For Each obj As Object In cmbParamGroup.Items
                        Dim it As ParamGroupItem = TryCast(obj, ParamGroupItem)
                        If it IsNot Nothing AndAlso it.Value = target Then
                            cmbParamGroup.SelectedItem = it
                            Exit For
                        End If
                    Next

                    If _param.Settings.AllowVaryBetweenGroups Then
                        rdoVaryEach.Checked = True
                    Else
                        rdoVaryAligned.Checked = True
                    End If
                End If

                ApplyBindingEnableState()
            End Sub

            Private Sub OnBindingTypeChanged(sender As Object, e As EventArgs)
                ApplyBindingEnableState()
            End Sub

            Private Sub ApplyBindingEnableState()
                Dim enableVary As Boolean = rdoInstance.Checked
                rdoVaryAligned.Enabled = enableVary
                rdoVaryEach.Enabled = enableVary
            End Sub

            Private Sub RebuildTree()
                Dim filter As String = Convert.ToString(cmbFilter.SelectedItem)
                If String.IsNullOrWhiteSpace(filter) Then filter = "<show all>"

                tvCats.BeginUpdate()
                tvCats.Nodes.Clear()

                _suppressTreeEvents = True
                For Each root As CategoryTreeItem In _categoryTree
                    Dim node As WinForms.TreeNode = BuildTreeNodeRecursive(root, filter, chkHideUnchecked.Checked)
                    If node IsNot Nothing Then tvCats.Nodes.Add(node)
                Next
                _suppressTreeEvents = False

                tvCats.EndUpdate()
            End Sub

            Private Function BuildTreeNodeRecursive(item As CategoryTreeItem, filter As String, hideUnchecked As Boolean) As WinForms.TreeNode
                If item Is Nothing Then Return Nothing

                Dim matchesFilter As Boolean = CategoryMatchesFilter(item, filter)
                Dim anyChild As Boolean = False

                Dim tn As New WinForms.TreeNode(item.Name)
                tn.Tag = item

                Dim isChecked As Boolean = False
                If _checked.ContainsKey(item.IdInt) Then isChecked = _checked(item.IdInt)
                tn.Checked = isChecked

                If Not item.IsBindable Then
                    tn.ForeColor = SystemColors.GrayText
                    tn.ToolTipText = item.Path & vbCrLf & "※ AllowsBoundParameters=False (Revit이 바인딩을 드랍할 수 있음)"
                Else
                    tn.ToolTipText = item.Path
                End If

                If item.Children IsNot Nothing AndAlso item.Children.Count > 0 Then
                    For Each ch As CategoryTreeItem In item.Children
                        Dim childNode As WinForms.TreeNode = BuildTreeNodeRecursive(ch, filter, hideUnchecked)
                        If childNode IsNot Nothing Then
                            tn.Nodes.Add(childNode)
                            anyChild = True
                        End If
                    Next
                End If

                If hideUnchecked Then
                    Dim keep As Boolean = tn.Checked OrElse HasAnyCheckedDescendant(tn)
                    If Not keep Then Return Nothing
                End If

                If Not matchesFilter AndAlso Not anyChild Then
                    Return Nothing
                End If

                Return tn
            End Function

            Private Function CategoryMatchesFilter(item As CategoryTreeItem, filter As String) As Boolean
                Select Case filter
                    Case "Model"
                        Return item.CatType = CategoryType.Model
                    Case "Annotation"
                        Return item.CatType = CategoryType.Annotation
                    Case "Analytical"
                        Return item.CatType = CategoryType.AnalyticalModel
                    Case "Internal"
                        Return item.CatType = CategoryType.Internal
                    Case Else
                        Return True
                End Select
            End Function

            Private Function HasAnyCheckedDescendant(node As WinForms.TreeNode) As Boolean
                For Each ch As WinForms.TreeNode In node.Nodes
                    If ch.Checked Then Return True
                    If HasAnyCheckedDescendant(ch) Then Return True
                Next
                Return False
            End Function

            Private Sub OnAfterCheckTree(sender As Object, e As WinForms.TreeViewEventArgs)
                If _suppressTreeEvents Then Return
                Dim it As CategoryTreeItem = TryCast(e.Node.Tag, CategoryTreeItem)
                If it Is Nothing Then Return

                Dim newState As Boolean = e.Node.Checked

                _suppressTreeEvents = True
                ApplyCheckToNodeAndDescendants(e.Node, newState)
                _suppressTreeEvents = False

                If chkHideUnchecked.Checked Then
                    RebuildTree()
                End If
            End Sub

            Private Sub ApplyCheckToNodeAndDescendants(node As WinForms.TreeNode, state As Boolean)
                If node Is Nothing Then Return

                Dim it As CategoryTreeItem = TryCast(node.Tag, CategoryTreeItem)
                If it IsNot Nothing Then
                    node.Checked = state
                    _checked(it.IdInt) = state
                End If

                For Each ch As WinForms.TreeNode In node.Nodes
                    ApplyCheckToNodeAndDescendants(ch, state)
                Next
            End Sub

            Private Sub SetAllChecked(value As Boolean)
                Dim allIds As New List(Of Integer)()
                CollectAllIds(_categoryTree, allIds)

                For Each idInt As Integer In allIds
                    _checked(idInt) = value
                Next

                RebuildTree()
            End Sub

            Private Sub CollectAllIds(items As List(Of CategoryTreeItem), ByRef outList As List(Of Integer))
                If items Is Nothing Then Return
                For Each it As CategoryTreeItem In items
                    If it Is Nothing Then Continue For
                    outList.Add(it.IdInt)
                    If it.Children IsNot Nothing AndAlso it.Children.Count > 0 Then
                        CollectAllIds(it.Children, outList)
                    End If
                Next
            End Sub

            Private Sub OnOk(sender As Object, e As EventArgs)
                Me.DialogResult = WinForms.DialogResult.OK
            End Sub

        End Class

#End Region

#Region "UI - Progress Form"

        Friend Class ProgressForm
            Inherits WinForms.Form

            Private lbl As WinForms.Label
            Private bar As WinForms.ProgressBar

            Public Sub New()
                Me.Text = "Processing..."
                Me.StartPosition = WinForms.FormStartPosition.CenterScreen
                Me.Width = 680
                Me.Height = 120
                Me.ControlBox = False
                Me.TopMost = True

                lbl = New WinForms.Label() With {.Dock = WinForms.DockStyle.Top, .Height = 35, .Text = "준비중...", .AutoEllipsis = True}
                bar = New WinForms.ProgressBar() With {.Dock = WinForms.DockStyle.Top, .Height = 20, .Minimum = 0, .Maximum = 100, .Value = 0}

                Me.Controls.Add(bar)
                Me.Controls.Add(lbl)
            End Sub

            Public Sub SetStatus(text As String)
                lbl.Text = text
                WinForms.Application.DoEvents()
            End Sub

            Public Sub SetProgress(current As Integer, total As Integer)
                If total <= 0 Then total = 1
                Dim v As Integer = CInt(Math.Truncate((CDbl(current) / CDbl(total)) * 100.0R))
                If v < 0 Then v = 0
                If v > 100 Then v = 100
                bar.Value = v
                WinForms.Application.DoEvents()
            End Sub
        End Class

#End Region

    End Class

End Namespace
