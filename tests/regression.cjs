const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const file = 'exp1_fcdt_traditional.js';
const js = fs.readFileSync(path.join(root, file), 'utf8');

function eventTarget(fields = {}) {
    const listeners = new Map();
    return Object.assign({
        addEventListener(type, handler) {
            if (!listeners.has(type)) listeners.set(type, new Set());
            listeners.get(type).add(handler);
        },
        removeEventListener(type, handler) { listeners.get(type)?.delete(handler); },
        async emit(type, event = {}) {
            for (const handler of [...(listeners.get(type) || [])]) {
                await handler({ preventDefault() {}, stopPropagation() {}, ...event });
            }
        }
    }, fields);
}

function load(entry = 'index.html') {
    const html = fs.readFileSync(path.join(root, entry), 'utf8');
    assert.match(html, /<script src="exp1_fcdt_traditional\.js\?v=20261126quality1"><\/script>/);
    const elements = {};
    const draw = new Proxy({}, { get: () => () => {} });
    for (const [, id] of html.matchAll(/\sid="([^"]+)"/g)) {
        elements[id] = eventTarget({
            value: '', classList: { add() {}, remove() {}, contains() {return false;} },
            style: {}, reportValidity: () => true, getContext: () => draw, focus() {}
        });
    }
    const timers = new Map();
    let nextTimer = 0;
    const alerts = [];
    const calls = [];
    const storage = Object.fromEntries(['initializeStorage', 'loadProgress', 'saveTrial', 'saveCheckpoint', 'saveProgress', 'saveFinalResult', 'clearProgress', 'finishExperiment'].map(name => [name, async (data) => {
        calls.push([name, data]);
        return name === 'initializeStorage' ? {backend: 'supabase'} : null;
    }]));
    const document = eventTarget({
        getElementById: id => elements[id] ?? null,
        hidden: false, hasFocus() { return !this.hidden; },
        documentElement: {}
    });
    const window = eventTarget({CLTStorage: storage, innerWidth: 1024, innerHeight: 768});
    const context = vm.createContext({
        window, document, history: {pushState() {}},
        HTMLInputElement: class {}, HTMLTextAreaElement: class {}, HTMLSelectElement: class {},
        setTimeout(callback, ms) { timers.set(++nextTimer, { callback, ms }); return nextTimer; },
        clearTimeout(handle) { timers.delete(handle); },
        performance: { now: () => 100 },
        alert: msg => alerts.push(msg),
        console: {log() {}, error() {}, warn() {}},
        XLSX: {utils: {json_to_sheet: rows => rows}}
    });
    vm.runInContext(js + `
        globalThis.hooks = {
            buildKRow, kSheetFromRows, buildFormalTrials, seedFromParticipantID, startFormal, finishExperiment,
            snapshot() { return { p, currentTrial, experimentPhase, stimData }; },
            testProbe() {
                currentTrial = 0; currentBlock = runtimeNumBlocks; trialData = { isChange: 1, block: currentBlock, setSize: 4, studyDuration: 250 };
                activeLocs = []; activeColors = [];
                showProbe({ change: 1 });
                return responseTimeoutTimer;
            },
            stubLocalExport(fn) { downloadXLSX = fn; },
            stubTrials(fn) { runTrial = fn; },
            endOneTrialRun() { currentBlock = runtimeNumBlocks; endBlock(); }
        };
    `, context, {filename: file});
    const a = context.hooks;
    return {a, elements, document, window, calls, alerts, timers};
}

const hit = { isChange: 1, response: 1, accuracy: 1, TrialStatus: 'valid', rt: 0.25, CountedInK: 1 };
const miss = { ...hit, response: 0, accuracy: 0 };
const correctReject = { ...hit, isChange: 0, response: 0 };
const falseAlarm = { ...correctReject, response: 1, accuracy: 0 };
const timeout = { isChange: 0, timeout: true, TrialStatus: 'timeout', CountedInK: 0 };
const interruption = { isChange: 0, timeout: false, TrialStatus: 'interrupted', CountedInK: 0 };

test('空分母和全虚报导出空白，真实零仍是数值', () => {
    const {a} = load();
    const row = a.buildKRow(4, 250, [hit, timeout, interruption]);
    assert.equal(row.HitRate, '1.0000');
    assert.equal(row.FalseAlarmRate, null);
    assert.equal(row.K, null);
    assert.equal(row.NNoChange, 0);
    assert.equal(row.ValidTrials, 1);
    assert.equal(row.TimeoutTrials, 1);
    assert.equal(row.InterruptedTrials, 1);
    assert.equal(row.TimeoutExclusionRate, '0.3333');
    const [sheetRow] = a.kSheetFromRows([row]);
    assert.ok(!Object.hasOwn(sheetRow, 'FalseAlarmRate'));
    assert.ok(!Object.hasOwn(sheetRow, 'K'));
    assert.equal(a.buildKRow(4, 250, [hit, falseAlarm]).K, null);
    const validZero = a.buildKRow(4, 250, [miss, correctReject]);
    assert.equal(a.kSheetFromRows([validZero])[0].K, 0);
    assert.equal(a.buildKRow(4, 250, [hit, correctReject]).K, '4.0000');
    const none = a.buildKRow(4, 250, [timeout, interruption]);
    for (const key of ['HitRate','FalseAlarmRate','K','Accuracy_Exclude']) assert.equal(none[key], null);
});

test('两个实际 HTML 入口均可开始，调试结束只导出本地文件', async () => {
    for (const name of ['index.html', 'exp1_fcdt_traditional.html']) {
        const {a, elements, document, window, calls} = load(name);
        elements.subGroup.value = '0';
        await elements['start-btn'].emit('click');
        assert.equal(a.snapshot().experimentPhase, 'instructions');
        assert.equal(a.snapshot().p.subjectID, '1111');
        a.testProbe();
        let exported = 0;
        a.stubLocalExport(() => { exported++; });
        document.hidden = true;
        await document.emit('visibilitychange');
        a.endOneTrialRun();
        assert.equal(a.snapshot().experimentPhase, 'ending_wait');
        assert.equal(elements['final-acc'].innerText, '本组暂无有效试次');
        await window.emit('keydown', {key: ' ', code: 'Space'});
        assert.equal(a.snapshot().experimentPhase, 'upload');
        assert.equal(exported, 1);
        assert.deepEqual(calls, []);
    }
});

test('探测中切走时停表、只记一次中断，并在恢复前台后才开始下一试次', async () => {
    const {a, elements, document, window, timers} = load();
    elements.subGroup.value = '0';
    await elements['start-btn'].emit('click');
    const timeoutID = a.testProbe();
    let advanced = 0;
    a.stubTrials(() => { advanced++; });
    document.hidden = true;
    await document.emit('visibilitychange');
    await window.emit('blur');
    assert.ok(!timers.has(timeoutID));
    const rows = a.snapshot().stimData;
    assert.equal(rows.length, 1);
    assert.equal(rows[0].TrialStatus, 'interrupted');
    assert.equal(rows[0].timeout, false);
    assert.equal(rows[0].rt, null);
    assert.equal(rows[0].CountedInK, 0);
    [...timers.values()].find(timer => timer.ms === 50).callback();
    assert.equal(advanced, 0);
    document.hidden = false;
    await document.emit('visibilitychange');
    await window.emit('focus');
    assert.equal(advanced, 1);
});

test('未失焦时真实超时口径仍保留', async () => {
    const {a, elements, timers} = load();
    elements.subGroup.value = '0';
    await elements['start-btn'].emit('click');
    const timeoutID = a.testProbe();
    timers.get(timeoutID).callback();
    const [row] = a.snapshot().stimData;
    assert.equal(row.TrialStatus, 'timeout');
    assert.equal(row.timeout, true);
    assert.equal(row.CountedInK, 0);
});

test('同一码跨会话可配对，完整种子与 80 个试次条件顺序可复现', async () => {
    const left = load(), right = load();
    for (const f of [left, right]) {
        f.elements.participantCode.value = 'Study_A1-Z';
        f.elements.subGroup.value = '1';
        await f.elements['start-btn'].emit('click');
        assert.equal(f.a.snapshot().p.subjectID, 'Study_A1-Z');
        assert.deepEqual(f.calls.map(([name]) => name), ['initializeStorage', 'loadProgress']);
    }
    const seed = left.a.snapshot().p.rndSeed;
    assert.equal(seed, right.a.snapshot().p.rndSeed);
    assert.notEqual(seed, left.a.seedFromParticipantID('Study_A1-x'));
    const sequence = JSON.stringify(left.a.buildFormalTrials());
    assert.equal(JSON.parse(sequence).length, 80);
    assert.equal(sequence, JSON.stringify(right.a.buildFormalTrials()));
    assert.notEqual(left.a.buildKRow(4, 250, []).SessionID, right.a.buildKRow(4, 250, []).SessionID);
});

test('错误格式不能开始；留空时保持随机匿名编号', async () => {
    const {a, elements, alerts} = load();
    elements.subGroup.value = '2';
    elements.participantCode.value = '中文';
    await elements['start-btn'].emit('click');
    assert.equal(a.snapshot().experimentPhase, 'setup');
    assert.equal(alerts.length, 1);
    elements.participantCode.value = '';
    await elements['start-btn'].emit('click');
    assert.match(a.snapshot().p.subjectID, /^2_\d{3}_[a-z\d]{3}$/);
});

test('保留 Z 和斜杠键映射；正式结束走原上传及本地备份', async () => {
    const yes = load(), no = load();
    for (const f of [yes, no]) {
        f.elements.subGroup.value = '1';
        await f.elements['start-btn'].emit('click');
        f.a.testProbe();
    }
    await yes.window.emit('keydown', {key: '/'});
    await no.window.emit('keydown', {key: 'Z'});
    assert.equal(yes.a.snapshot().stimData[0].response, 1);
    assert.equal(yes.a.snapshot().stimData[0].accuracy, 1);
    assert.equal(no.a.snapshot().stimData[0].response, 0);
    assert.equal(no.a.snapshot().stimData[0].accuracy, 0);
    let exported = 0;
    yes.a.stubLocalExport(() => { exported++; });
    yes.a.endOneTrialRun();
    assert.equal(yes.a.snapshot().experimentPhase, 'ending_wait');
    assert.match(yes.elements['final-acc'].innerText, /100%/);
    await yes.window.emit('keydown', {key: ' ', code: 'Space'});
    for (let i = 0; i < 12; i++) await Promise.resolve();
    assert.equal(exported, 1);
    assert.deepEqual(yes.calls.map(([name]) => name), [
        'initializeStorage', 'loadProgress', 'saveTrial', 'saveCheckpoint',
        'saveFinalResult', 'clearProgress', 'finishExperiment'
    ]);
});
