# KKY Tool Revit — UI Patch (fix2 규격 반영)

## 구조
- /UI/Hub/HubHostWindow.vb
- /UI/Hub/UiBridgeExternalEvent.vb
- /Resources/HubUI/index.html
- /Resources/HubUI/main.js
- /Resources/HubUI/styles.css

## 적용 방법
1) 위 파일/폴더를 프로젝트 **동일 경로**로 복사(경로 없으면 생성) 후 덮어쓰기.
2) 참조 확인:
   - Microsoft.Web.WebView2 (WPF)
   - NPOI (2.6+), NPOI.OOXML
   - System.Windows.Forms (FolderBrowserDialog)
3) 허브에서 기능 진입 시 JS 호출:
   - 중복검토: `window.UI.renderDuplicateInspector()`
   - 커넥터: `window.UI.renderConnectorDiagnostics()`
   - 좌표/북각: `window.UI.renderExportPoints()`
4) Revit AddIn 명령에서: `New HubHostWindow(uiapp).Show()`

## 이벤트 매핑 (fix2)
- 중복검토: `dup:run`, `duplicate:export`, `duplicate:delete`, `duplicate:restore`, `duplicate:select`
- 커넥터: `connector:run`, `connector:save-excel`
- 좌표/북각: `export:browse-folder`, `export:preview`, `export:save-excel`
- 공유 파라미터 연동: `sharedparam:list`, `sharedparam:run`, `sharedparam:export-excel` (호환: `paramprop:run`)

## 엑셀 스키마 (고정)
- 중복검토: `그룹# | ElementId | 카테고리 | 패밀리 | 타입 | 연결수(실 연결) | 연결 객체 | 삭제 후보`
- 커넥터: `Id1 | Id2 | Category1 | Category2 | Family1 | Family2 | Distance (inch) | ConnectionType | ParamName | Value1 | Value2 | Status`
- 좌표/북각: `File | ProjectPoint_E(mm) | ProjectPoint_N(mm) | ProjectPoint_Z(mm) | SurveyPoint_E(mm) | SurveyPoint_N(mm) | SurveyPoint_Z(mm) | TrueNorthAngle(deg)`
