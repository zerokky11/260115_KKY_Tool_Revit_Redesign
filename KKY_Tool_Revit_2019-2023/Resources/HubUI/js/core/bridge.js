// Resources/HubUI/js/core/bridge.js
// 공개 API: post, onHost, setConn, ping
// - WebView2(Revit) : window.chrome.webview.postMessage
// - DEV(브라우저 단독) : 간단 목업 동작
//
// 변경점
// 1) 웹→호스트 송신을 { ev, payload, __seq } 로 교정  ←★ 중요
// 2) 상세 로깅 강화: in/out 이벤트, payload 크기(B), 에러 message/stack 출력
// 3) onHost 두 형태 지원: onHost('ev', handler) + onHost((msg)=>{ev,payload})
// 4) 왕복 시간(ms) 표기: 호스트가 __seq를 되돌려주면 측정됨

import { log } from './dom.js';

export let DEV = false;

const hasWebView =
    typeof window !== "undefined" &&
    window.chrome &&
    window.chrome.webview &&
    typeof window.chrome.webview.postMessage === "function";

DEV = !hasWebView;

// -----------------------------
// 내부: 리스너 풀
// -----------------------------
const _byName = new Map();   // ev -> Set<fn(payload)>
const _generic = new Set();  // Set<fn({ ev, payload })>

// -----------------------------
// 상세 로깅 / 측정
// -----------------------------
let VERBOSE = true;
let __seq = 0;
const __pending = new Map(); // seq -> { ev, t0 }

function _sizeOf(obj) {
    try { return new Blob([JSON.stringify(obj)]).size; } catch { return 0; }
}
function _logIn(ev, payload, extra = "") {
    if (!VERBOSE) return;
    const sz = _sizeOf(payload);
    log(`[bridge] ← in ${ev}${extra} (${sz}B)`, payload);
}
function _logOut(ev, payload) {
    if (!VERBOSE) return;
    const sz = _sizeOf(payload);
    log(`[bridge] → out ${ev} (${sz}B)`, payload);
}
function _logErr(ctx, err) {
    const msg = err && err.message ? err.message : String(err);
    const stk = err && err.stack ? `\n${err.stack}` : '';
    console.error(`[bridge] ${ctx} error: ${msg}${stk}`, err);
}

// Shift+F12 → 상세로그 토글
window.addEventListener('keydown', (e) => {
    if (e.key === 'F12' && e.shiftKey) {
        VERBOSE = !VERBOSE;
        console.log(`[bridge] verbose ${VERBOSE ? "ON" : "OFF"}`);
    }
});

// -----------------------------
// 내부: 호스트 메시지 브로드캐스트
// -----------------------------
function _emitHost(ev, payload, meta) {
    // 왕복 시간 표기(호스트가 __seq를 반사하는 경우)
    let extra = "";
    const seq = meta?.__seq;
    if (seq && __pending.has(seq)) {
        const t0 = __pending.get(seq).t0;
        const dt = (performance.now() - t0).toFixed(1);
        extra = ` [${dt}ms,#${seq}]`;
        __pending.delete(seq);
    }
    _logIn(ev, payload, extra);

    // 이름 기반
    const set = _byName.get(ev);
    if (set && set.size) {
        for (const fn of set) {
            try { fn(payload); }
            catch (e) { _logErr(`onHost("${ev}") handler`, e); }
        }
    }
    // 제네릭 기반
    if (_generic.size) {
        for (const fn of _generic) {
            try { fn({ ev, payload }); }
            catch (e) { _logErr("onHost(generic) handler", e); }
        }
    }
}

// -----------------------------
// 공개: onHost(handler) | onHost('ev', handler)
// -----------------------------
export function onHost(arg1, arg2) {
    // onHost((msg)=>{})
    if (typeof arg1 === "function" && !arg2) {
        _generic.add(arg1);
        return () => _generic.delete(arg1);
    }
    // onHost('ev', handler)
    const ev = String(arg1 || "");
    const handler = arg2;
    if (!ev || typeof handler !== "function") return () => { };
    if (!_byName.has(ev)) _byName.set(ev, new Set());
    _byName.get(ev).add(handler);
    return () => _byName.get(ev)?.delete(handler);
}

// -----------------------------
// 공개: post(ev, payload) — ★ 호스트로 ev 키로 보냄
// -----------------------------
export function post(ev, payload = {}) {
    _logOut(ev, payload);

    const seq = ++__seq;
    __pending.set(seq, { ev, t0: performance.now() });

    if (!DEV) {
        try {
            // 호스트가 __seq를 그대로 실어 응답(ev/payload)와 함께 돌려주면 왕복시간 측정 가능
            window.chrome.webview.postMessage({ ev, payload, __seq: seq });
        } catch (e) {
            _logErr(`post("${ev}")`, e);
        }
        return;
    }

    // === DEV 목업 ===
    switch (ev) {
        case "ui:toggle-topmost": {
            __devTopMost = !__devTopMost;
            _emitHost("host:topmost", { on: __devTopMost }, { __seq: seq });
            break;
        }
        case "dup:run": {
            const rows = [
                { id: "12345", category: "Pipes", family: "Pipe Fitting", type: "Elbow - 90deg", groupId: 1, connectedCount: 2, connectedIds: "12346,12347", candidate: true },
                { id: "20011", category: "Pipes", family: "Pipe", type: "DN50", groupId: 1, connectedCount: 1, connectedIds: "20012", candidate: false },
            ];
            _emitHost("dup:list", { data: rows }, { __seq: seq });
            break;
        }
        case "duplicate:export": {
            _emitHost("dup:exported", { path: "C:\\Temp\\duplicate.xlsx" }, { __seq: seq });
            break;
        }
        case "connector:run": {
            const rows = [
                {
                    Id1: 101, Id2: 102, Category1: "Pipe", Category2: "Pipe", Family1: "Pipe", Family2: "Pipe",
                    "Distance (inch)": 0.12, ConnectionType: "Butt", ParamName: "System Type",
                    Value1: "Domestic Cold Water", Value2: "Domestic Cold Water", Status: "OK"
                },
            ];
            _emitHost("connector:done", { rows }, { __seq: seq });
            break;
        }
        case "export:preview": {
            const rows = [
                {
                    File: "Sample_A.rvt", "ProjectPoint_E(mm)": 100.0, "ProjectPoint_N(mm)": 200.0, "ProjectPoint_Z(mm)": 0.0,
                    "SurveyPoint_E(mm)": 110.0, "SurveyPoint_N(mm)": 210.0, "SurveyPoint_Z(mm)": 5.0, "TrueNorthAngle(deg)": 23.5
                },
            ];
            _emitHost("export:previewed", { rows }, { __seq: seq });
            break;
        }
        case "export:save-excel": {
            _emitHost("export:saved", { path: "C:\\Temp\\points.xlsx" }, { __seq: seq });
            break;
        }
        case "paramprop:run": {
            _emitHost("paramprop:done", { ok: true, status: "succeeded", message: "(DEV) 완료" }, { __seq: seq });
            _emitHost("paramprop:report", { report: "(DEV) 완료" }, { __seq: seq });
            break;
        }
        case "familylink:init": {
            _emitHost("familylink:sharedparams", {
                sourcePath: "C:\\Sample\\SharedParameters.txt",
                items: [
                    { name: "PARAM_A", guid: "11111111-1111-1111-1111-111111111111", groupName: "GroupA", dataTypeToken: "Text" },
                    { name: "PARAM_B", guid: "22222222-2222-2222-2222-222222222222", groupName: "GroupB", dataTypeToken: "Length" }
                ]
            }, { __seq: seq });
            break;
        }
        case "familylink:pick-rvts": {
            _emitHost("familylink:rvts-picked", { paths: ["C:\\Sample\\ProjectA.rvt", "C:\\Sample\\ProjectB.rvt"] }, { __seq: seq });
            break;
        }
        case "familylink:run": {
            _emitHost("familylink:progress", { percent: 25, message: "(DEV) 스캔 중..." }, { __seq: seq });
            _emitHost("familylink:result", {
                schema: ["ProjectPath", "HostFamilyName", "HostFamilyCategory", "NestedFamilyName", "NestedTypeName", "NestedCategory", "TargetParamName", "ExpectedGuid", "FoundScope", "NestedParamGuid", "NestedParamDataType", "AssocHostParamName", "HostParamGuid", "HostParamIsShared", "Issue", "Notes"],
                rows: [
                    {
                        ProjectPath: "C:\\Sample\\ProjectA.rvt",
                        HostFamilyName: "HostA",
                        HostFamilyCategory: "Generic Model",
                        NestedFamilyName: "NestedA",
                        NestedTypeName: "Type01",
                        NestedCategory: "Generic Model",
                        TargetParamName: "PARAM_A",
                        ExpectedGuid: "11111111-1111-1111-1111-111111111111",
                        FoundScope: "InstanceParam",
                        NestedParamGuid: "11111111-1111-1111-1111-111111111111",
                        NestedParamDataType: "Text",
                        AssocHostParamName: "PARAM_A",
                        HostParamGuid: "11111111-1111-1111-1111-111111111111",
                        HostParamIsShared: "True",
                        Issue: "OK",
                        Notes: ""
                    }
                ]
            }, { __seq: seq });
            break;
        }
        case "familylink:export": {
            _emitHost("familylink:exported", { ok: true, path: "C:\\Temp\\FamilyLinkAudit.xlsx" }, { __seq: seq });
            break;
        }
        default: {
            console.debug("[DEV] post noop:", ev, payload);
            break;
        }
    }
}

// -----------------------------
// 공개: 연결 상태 칩
// -----------------------------
export function setConn(ok) {
    const chip = document.querySelector('#connChip');
    if (!chip) return;
    chip.classList.toggle('is-active', !!ok);
    chip.classList.toggle('is-disconnected', !ok);
    const label = chip.querySelector('.chip-text');
    if (label) label.textContent = ok ? '연결됨' : '연결 끊김';
}

// -----------------------------
// 공개: ping
// -----------------------------
export function ping() {
    if (!DEV) { setConn(true); return; }
    setConn(false);
    setTimeout(() => setConn(true), 350);
}

// -----------------------------
// 초기: 호스트 메시지 수신(WebView2)
// -----------------------------
if (!DEV) {
    window.chrome.webview.addEventListener("message", (e) => {
        const data = e?.data;
        if (!data) return;
        const ev = data.ev || data.name || "";  // 호스트가 ev 또는 name 둘 다 가능
        _emitHost(ev, data.payload, { __seq: data.__seq });
    });
}

// -----------------------------
// DEV 초기 상태
// -----------------------------
let __devTopMost = false;
