// Run with Node.js and Playwright: node tests/mention-form-attendance.cjs
// ActiveX file writes and Worker launch are mocked; this does not send real requests.
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const { chromium } = require(process.env.CODEX_PRIMARY_RUNTIME_NODE_MODULES
    ? path.join(process.env.CODEX_PRIMARY_RUNTIME_NODE_MODULES, 'playwright') : 'playwright');

const folder = path.join(__dirname, '..', 'メンション依頼フォーム_HTA_最新版');
const html = fs.readFileSync(path.join(folder, 'メンション依頼フォーム.hta'), 'utf8')
    .replace(/^\uFEFF/, '').replace(/<hta:application\b[\s\S]*?\/>/i, '')
    .replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, '')
    .replace(/<link\b[^>]*>/gi, '').replace(' onload="initApp();"', '');
// Strip the transport BOM when injecting CSS as text, as external CSS loading does.
const css = fs.readFileSync(path.join(folder, 'mention-form.css'), 'utf8').replace(/^\uFEFF/, '');
const source = fs.readFileSync(path.join(folder, 'mention-form.js'), 'utf8');

(async () => {
    const browser = await chromium.launch({ headless: true, executablePath: process.env.MENTION_TEST_BROWSER || undefined });
    try {
        const page = await browser.newPage({ viewport: { width: 1530, height: 1000 } });
        const errors = [];
        page.on('pageerror', e => errors.push(e.message));
        await page.setContent(html);
        await page.addStyleTag({ content: css });
        if (process.env.MENTION_TEST_FONT) {
            const font = fs.readFileSync(process.env.MENTION_TEST_FONT).toString('base64');
            await page.addStyleTag({ content: `@font-face{font-family:TestJapanese;src:url(data:font/ttf;base64,${font})}body,input,select,button{font-family:TestJapanese,sans-serif!important}` });
            await page.evaluate(() => document.fonts.ready);
        }
        await page.addScriptTag({ content: source });
        const results = await page.evaluate(() => {
            var passed=[];
            function ok(condition,name){ if(!condition){ throw new Error(name); } passed.push(name); }
            function fire(id){ $(id).dispatchEvent(new Event('input',{bubbles:true})); }
            function select(no,status){ $('attendance'+status+no).click(); }
            function shown(id){ return getComputedStyle($(id)).display!=='none' && getComputedStyle($(id)).visibility!=='hidden'; }
            function fill(no,status){
                $('org'+no).value='組織'+no; $('ca'+no).value='テストCA'+no;
                $('mailMemo'+no).value='テストメモ'+no; $('processedAt'+no).value='2026/09/10 15:00';
                $('type'+no).selectedIndex=1;
                if(status){ select(no,status); }
            }
            function reset(count){
                for(var n=1;n<=3;n++){
                    clearRequestFields(n);
                    $('request'+n).className='request-card card'+n+(n>count?' hidden':'');
                }
                visibleRequestCount=count;
                closeValidationModal(); closeSendConfirm(); hideSystemModal();
                $('requestBase').value='呉服';
                saved=[]; launched=[]; accepted=0;
            }
            function snapshot(no){
                var r=collectRequest(no);
                return JSON.stringify([r.organization,r.caName,r.attendance,r.proxyCA1,r.proxyCA2,r.proxyCA3,r.urgent,r.noProxy]);
            }
            populateBases(); populateTypes(); bindSameCASync();
            // Only Windows and window-management boundaries are replaced.
            resizeApp=function(){}; hideAppForLayout=function(){};
            $('app').style.visibility='visible';
            $('requesterName').innerText='動作確認用';
            var saved=[], launched=[], accepted=0;
            getProductionCsvFolder=function(){return 'TEST_ONLY';};
            savePendingPackage=function(base,requester,id,sentAt,requests,lines){
                saved.push({requests:requests,lines:lines}); return 'TEST_PENDING';
            };
            removeCompletionReadyMarkerIfExists=function(){};
            beginPendingState=function(){};
            launchBackgroundWorker=function(){launched.push(true);};
            showAcceptedAndPrepareMinimize=function(){accepted++;};
            reset(1);
            ok(!getAttendance(1) && shown('options1') && !shown('urgentOption1') && !shown('noProxyOption1'),'initial status unset, option heading visible and choices hidden');
            fill(1);
            ok(!validateForm() && modalFocusTarget==='attendanceWork1','attendance required before mail fields');
            closeValidationModal();
            select(1,'Work');
            ok(shown('options1') && shown('urgentOption1') && !shown('noProxyOption1'),'work shows urgent only');
            ok($('proxyCA1_1').disabled && $('proxyCA1_2').disabled && $('proxyCA1_3').disabled && validateForm(),'work accepts blank disabled proxies');
            $('urgent1').click();
            ok(collectRequest(1).urgent,'urgent selected for working CA');
            select(1,'Off');
            ok(!checked('urgent1') && !shown('urgentOption1') && shown('noProxyOption1'),'off clears urgent and shows no-proxy');
            ok(!$('proxyCA1_1').disabled && !validateForm() && modalFocusTarget==='proxyCA1_1','off requires at least one proxy');
            closeValidationModal();
            $('proxyCA1_1').value='　  ';
            ok(!validateForm(),'whitespace proxy is rejected'); closeValidationModal();
            $('proxyCA1_1').value=''; $('proxyCA1_2').value='代理2';
            ok(validateForm(),'a proxy in the second row is accepted');
            $('noProxy1').click();
            ok(validateForm() && !$('proxyCA1_2').value && $('proxyCA1_2').disabled && collectRequest(1).noProxy,'no-proxy waives requirement and clears names');
            $('noProxy1').click();
            ok(!$('proxyCA1_1').disabled && !validateForm(),'unchecking no-proxy restores requirement'); closeValidationModal();
            $('proxyCA1_3').value='代理3'; select(1,'Short');
            ok(validateForm() && collectRequest(1).shortTime && !collectRequest(1).urgent,'short-time status uses existing short-time flag');
            ok($('proxyCA1_3').value==='代理3','off to short keeps a valid proxy');
            select(1,'Work');
            ok(!$('proxyCA1_3').value && !checked('noProxy1') && !collectRequest(1).shortTime,'switch to work clears proxy and short-time flag');

            // Validation sequence follows the visible left-to-right fields, then the footer.
            reset(1);
            ok(!validateForm() && modalFocusTarget==='org1','organization is first required field'); closeValidationModal();
            $('org1').value='組織';
            ok(!validateForm() && modalFocusTarget==='ca1','CA is second required field'); closeValidationModal();
            $('ca1').value='CA'; select(1,'Work');
            ok(!validateForm() && modalFocusTarget==='mailMemo1','mail follows CA state'); closeValidationModal();
            $('mailMemo1').value='メモ';
            ok(!validateForm() && modalFocusTarget==='processedAt1','processed time follows mail'); closeValidationModal();
            $('processedAt1').value='日時';
            ok(!validateForm() && modalFocusTarget==='type1','type follows time'); closeValidationModal();
            $('type1').selectedIndex=1; $('requestBase').value='';
            ok(!validateForm() && modalFocusTarget==='requestBase','base is last required field'); closeValidationModal();

            reset(3); fill(1,'Off'); fill(2,'Work'); fill(3,'Work');
            $('proxyCA1_1').value='代理A'; $('proxyCA1_2').value='代理B'; $('proxyCA1_3').value='代理C';
            $('sameCA3').click(); $('sameCA2').click();
            ok(snapshot(1)===snapshot(2) && snapshot(2)===snapshot(3),'linking request 2 after request 3 updates all CA fields');
            ok($('org2').disabled && $('attendanceOff2').disabled && $('noProxy2').disabled && $('proxyCA3_3').disabled,'inherited CA, radios, options and proxies are locked');
            ok($('mailMemo2').value==='テストメモ2' && !$('mailMemo2').disabled,'mail and other request details stay independent');
            ok(!shown('inputNote2') && !shown('inputNote3'),'same-CA notes hidden');
            $('ca1').value='変更CA'; fire('ca1'); $('proxyCA1_3').value='変更代理'; fire('proxyCA1_3');
            ok($('ca3').value==='変更CA' && $('proxyCA3_3').value==='変更代理','text input propagates through both links');
            $('noProxy1').click();
            ok(checked('noProxy3') && !$('proxyCA3_1').value && snapshot(1)===snapshot(3),'no-proxy selection propagates and clears all linked proxies');
            select(1,'Work'); $('urgent1').click();
            ok(checked('urgent3') && shown('urgentOption3') && !shown('noProxyOption3') && snapshot(1)===snapshot(3),'attendance and urgent propagate with matching visibility');
            $('sameCA2').click();
            ok(!$('org2').value && !getAttendance(2) && shown('options2') && !checked('urgent2'),'unlink clears organization and choices while keeping option heading');
            ok(!$('org3').value && !getAttendance(3) && shown('options3'),'unlink clears downstream values while keeping option heading');
            ok(!$('org2').disabled && !$('attendanceWork2').disabled && shown('inputNote2'),'unlink unlocks independent input and restores note');
            fill(2,'Short'); $('noProxy2').click();
            ok(snapshot(2)===snapshot(3),'new independent state propagates to request 3');
            // Programmatic changes that did not fire input still sync before collection.
            $('ca2').value='送信直前CA';
            ok(validateForm() && $('ca3').value==='送信直前CA','validation synchronizes before submission');
            clearRequest(2);
            ok(!getAttendance(2) && !getAttendance(3) && shown('options3'),'clear resets linked attendance and choices while keeping option heading');

            // All four cancellation relationships: independent/linked request 2 and 3.
            for(var link2=0;link2<2;link2++){
                for(var link3=0;link3<2;link3++){
                    reset(3); fill(1,'Work'); $('urgent1').click();
                    fill(2,'Off'); $('noProxy2').click(); fill(3,'Short'); $('proxyCA3_1').value='個別代理';
                    if(link2){ $('sameCA2').click(); }
                    if(link3){ $('sameCA3').click(); }
                    var original=snapshot(3), memo=$('mailMemo3').value;
                    cancelRequest(2);
                    ok(snapshot(2)===original && $('mailMemo2').value===memo,'promotion preserves values '+link2+link3);
                    ok(checked('sameCA2')===!!(link2&&link3) && !checked('sameCA3') && !getAttendance(3),'promotion resets obsolete links '+link2+link3);
                    $('ca1').value='取消後変更'; fire('ca1');
                    ok(link2&&link3 ? $('ca2').value==='取消後変更' : snapshot(2)===original,'only a genuine prior chain remains linked '+link2+link3);
                }
            }
            reset(3); fill(1,'Off'); $('noProxy1').click(); $('sameCA2').click(); $('sameCA3').click();
            cancelRequest(3);
            ok(!checked('sameCA3') && !getAttendance(3) && shown('options3'),'cancel request 3 resets its fields without hiding option heading');
            resetAfterSend();
            ok(visibleRequestCount===1 && !getAttendance(1) && !getAttendance(2) && !getAttendance(3) && !checked('sameCA2') && !checked('sameCA3'),'post-send reset clears all inherited state');
            ok($('requestBase').value==='呉服' && $('requesterName').innerText==='動作確認用','post-send reset preserves base and requester');

            reset(1); fill(1,'Short'); $('noProxy1').click();
            var before=JSON.stringify(collectRequest(1));
            sendRequest();
            ok(sendConfirmationOpen && shown('sendConfirmOverlay') && !saved.length && !launched.length,'send opens confirmation without writing Pending or launching Worker');
            ok($('sendConfirmMessage').innerText==='出社状況は相違ないですか？' && document.activeElement===$('sendConfirmNo'),'confirmation wording and initial focus');
            $('sendConfirmNo').click();
            ok(!sendConfirmationOpen && !shown('sendConfirmOverlay') && before===JSON.stringify(collectRequest(1)) && !saved.length,'No preserves form and performs no send');
            sendRequest(); $('sendConfirmYes').click();
            ok(saved.length===1 && launched.length===1 && accepted===1 && !sendConfirmationOpen,'Yes invokes existing send pipeline once');
            ok(saved[0].requests[0].shortTime && !saved[0].requests[0].proxyCA1,'confirmed request uses current status');
            confirmSendRequest();
            ok(saved.length===1,'dismissed confirmation cannot submit again');
            reset(1); fill(1);
            sendRequest();
            ok(!sendConfirmationOpen && !saved.length && modalFocusTarget==='attendanceWork1','invalid attendance never opens send confirmation'); closeValidationModal();

            // Parse generated CSV with the unchanged Worker's actual parser and header checker.
            return passed;
        });
        // The real Worker parser is exercised without evaluating the top-level Windows I/O.
        const worker = fs.readFileSync(path.join(folder, 'mention-request-worker.js'), 'utf8');
        const workerFunctions = worker.slice(worker.indexOf('    function validateHeader('), worker.lastIndexOf('})();'));
        const header = worker.match(/var HEADER = \[[\s\S]*?\];/)[0];
        await page.addScriptTag({ content: header + '\n' + workerFunctions });
        const csvChecks = await page.evaluate(() => {
            var records=[];
            for(var status of ['Work','Off','Short']){
                clearRequestFields(1); $('org1').value='組織'; $('ca1').value='CA';
                $('attendance'+status+'1').click();
                if(status==='Work'){ $('urgent1').click(); } else { $('noProxy1').click(); }
                $('mailMemo1').value='引用"と,改行\nメモ';
                var text=makeHeaderLine()+'\r\n'+makeCsvLine('TEST','2026/09/10 15:00:00','呉服','確認',collectRequest(1));
                var rows=parseCsv(text); validateHeader(rows[0]);
                if(rows[1].length!==17 || rows[1][7]!=='' || rows[1][11]!==$('mailMemo1').value || rows[1][14]!== (status==='Short'?'●':'') || rows[1][15]!== (status==='Work'?'●':'')){
                    throw new Error('CSV compatibility failed: '+status);
                }
                records.push('CSV schema and existing flags: '+status);
            }
            return records;
        });
        results.push(...csvChecks);
        await page.evaluate(() => {
            closeValidationModal(); visibleRequestCount=3;
            for(var n=1;n<=3;n++){
                clearRequestFields(n); $('request'+n).className='request-card card'+n;
                $('org'+n).value='サンプル組織'; $('ca'+n).value='サンプルCA';
                $('mailMemo'+n).value='メールメモ'; $('processedAt'+n).value='2026/09/10 15:00';
                $('dueDate'+n).value='2026/09/11'; $('type'+n).selectedIndex=5;
            }
            $('attendanceWork1').click(); $('urgent1').click();
            $('attendanceOff2').click(); $('proxyCA2_1').value='代理CA';
            $('attendanceShort3').click(); $('noProxy3').click();
        });
        const layout = await page.evaluate(() => {
            function rect(el){return el.getBoundingClientRect();}
            var failures=[];
            for(var n=1;n<=3;n++){
                var card=$('request'+n), main=card.querySelector('.main-row');
                var blocks=Array.from(main.children);
                for(var i=1;i<blocks.length;i++){
                    if(rect(blocks[i]).left<rect(blocks[i-1]).right-1 || Math.abs(rect(blocks[i]).top-rect(blocks[0]).top)>1){ failures.push('column overlap or wrap '+n+':'+i); }
                }
                var detail=rect(card.querySelector('.detail-zone'));
                var row=rect(card.querySelector('.card-action-row'));
                var buttons=card.querySelectorAll('.card-action-row button');
                if(Math.abs(detail.right-rect(buttons[buttons.length-1]).right)>1 || Math.abs(detail.left-row.left)>1){ failures.push('button alignment '+n); }
                if(rect($('proxyCA'+n+'_3')).bottom>rect(card).bottom || row.bottom>rect(card).bottom){ failures.push('vertical overflow '+n+': proxy='+rect($('proxyCA'+n+'_3')).bottom+', buttons='+row.bottom+', card='+rect(card).bottom); }
                for(var el of card.querySelectorAll('label')){
                    if(getComputedStyle(el).display!=='none' && el.scrollWidth>el.clientWidth+1){ failures.push('label clipped '+n+':'+el.innerText); }
                }
            }
            return failures;
        });
        if(process.env.MENTION_TEST_OUTPUT){
            fs.mkdirSync(process.env.MENTION_TEST_OUTPUT,{recursive:true});
            await page.locator('#app').screenshot({path:path.join(process.env.MENTION_TEST_OUTPUT,'attendance-form.png')});
        }
        assert.deepEqual(layout, [], 'layout bounds');
        results.push('three-card layout: columns, labels, buttons and vertical bounds');
        await page.click('#sendButton');
        assert.equal(await page.locator('#sendConfirmOverlay').isVisible(), true);
        await page.keyboard.press('Tab');
        assert.equal(await page.evaluate(()=>document.activeElement.id), 'sendConfirmYes');
        await page.keyboard.press('Tab');
        assert.equal(await page.evaluate(()=>document.activeElement.id), 'sendConfirmNo');
        results.push('keyboard focus stays inside send confirmation');
        if(process.env.MENTION_TEST_OUTPUT){
            await page.screenshot({path:path.join(process.env.MENTION_TEST_OUTPUT,'attendance-confirmation.png')});
        }
        await page.keyboard.press('Escape');
        assert.equal(await page.locator('#sendConfirmOverlay').isVisible(), false);
        assert.equal(await page.inputValue('#ca1'), 'サンプルCA');
        results.push('Escape returns to the unchanged form');
        assert.deepEqual(errors, [], 'browser JavaScript errors');
        console.log(JSON.stringify({passed:results.length,checks:results}, null, 2));
    } finally {
        await browser.close();
    }
})().catch(e => { console.error(e); process.exitCode=1; });
