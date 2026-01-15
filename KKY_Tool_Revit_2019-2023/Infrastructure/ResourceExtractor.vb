Imports System.IO
Imports System.IO.Compression
Imports System.Reflection
Imports System.Security.Cryptography
Imports System.Text
Imports System.Linq

' RootNamespace: KKY_Tool_Revit
Public NotInheritable Class ResourceExtractor
    Private Sub New()
    End Sub

    ''' <summary>
    ''' HubUI.zip을 추출(또는 캐시 유지)하고 targetDir를 반환합니다.
    ''' </summary>
    Public Shared Sub EnsureExtractedUI(targetDir As String)
        Directory.CreateDirectory(targetDir)

        Dim asm = Assembly.GetExecutingAssembly()
        Dim resNames = asm.GetManifestResourceNames()

        ' 1) 임베디드 리소스에서 찾기
        Dim resName As String = resNames.FirstOrDefault(Function(n) n.EndsWith("HubUI.zip", StringComparison.OrdinalIgnoreCase))
        If String.IsNullOrEmpty(resName) Then
            resName = resNames.FirstOrDefault(Function(n) n.ToLowerInvariant().Contains("hubui") AndAlso n.ToLowerInvariant().EndsWith(".zip"))
        End If

        Dim stream As Stream = Nothing
        If Not String.IsNullOrEmpty(resName) Then
            stream = asm.GetManifestResourceStream(resName)
        End If

        ' 2) 폴더 배포 대비: DLL 옆에서 HubUI.zip 찾기 (Resources\HubUI.zip, HubUI.zip)
        Dim asmDir = Path.GetDirectoryName(asm.Location)
        Dim diskCandidates = New String() {
          Path.Combine(asmDir, "Resources", "HubUI.zip"),
          Path.Combine(asmDir, "HubUI.zip")
        }
        If stream Is Nothing Then
            For Each p In diskCandidates
                If File.Exists(p) Then
                    stream = File.OpenRead(p)
                    Exit For
                End If
            Next
        End If

        If stream Is Nothing Then
            Dim msg As New StringBuilder()
            msg.AppendLine("임베디드 리소스 'HubUI.zip'을 찾지 못했습니다.")
            msg.AppendLine("어셈블리: " & asm.Location)
            msg.AppendLine("임베디드 목록(최대 20개):")
            For Each n In resNames.Take(20)
                msg.AppendLine(" - " & n)
            Next
            msg.AppendLine("디스크에서 시도한 경로:")
            For Each p In diskCandidates
                msg.AppendLine(" - " & p)
            Next
            Throw New FileNotFoundException(msg.ToString())
        End If

        Using s = stream
            Dim newHash = Sha256Hex(s)
            Dim stamp = Path.Combine(targetDir, ".ui_hash")

            Dim need As Boolean = True
            If File.Exists(stamp) Then
                Try
                    If String.Equals(File.ReadAllText(stamp, Encoding.UTF8), newHash, StringComparison.OrdinalIgnoreCase) Then
                        need = False
                    End If
                Catch
                End Try
            End If

            If need Then
                ' 기존 폴더 비우기
                Try
                    For Each f In Directory.EnumerateFiles(targetDir, "*", SearchOption.AllDirectories)
                        Try : File.SetAttributes(f, FileAttributes.Normal) : Catch : End Try
                    Next
                    Directory.Delete(targetDir, True)
                Catch
                End Try
                Directory.CreateDirectory(targetDir)

                ' ZIP 추출 (.NET FW는 overwrite 매개변수 없음)
                s.Position = 0
                Using z = New ZipArchive(s, ZipArchiveMode.Read, leaveOpen:=False)
                    z.ExtractToDirectory(targetDir)
                End Using

                File.WriteAllText(stamp, newHash, Encoding.UTF8)
            End If
        End Using
    End Sub

    Private Shared Function Sha256Hex(stream As Stream) As String
        stream.Position = 0
        Using sha = SHA256.Create()
            Return BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "").ToLowerInvariant()
        End Using
    End Function

End Class
