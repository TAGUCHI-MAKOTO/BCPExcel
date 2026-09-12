from pathlib import Path
import hashlib

root=Path('.')
js_path=root/'メンション依頼フォーム_HTA_最新版'/'mention-form.js'
css_path=root/'メンション依頼フォーム_HTA_最新版'/'mention-form.css'
readme_path=root/'メンション依頼フォーム_HTA_最新版'/'README.txt'

def replace_once(text, old, new, label):
    if old not in text:
        raise SystemExit('PATCH_MISS: '+label)
    return text.replace(old,new,1)

# CSS v28.23 -> v28.27
css=css_path.read_text(encoding='utf-8-sig')
css=replace_once(css,
'''.option-zone{\n    width:148px;\n    margin-left:0;\n    margin-right:12px;\n}\n\n/* 左側は (160 + 12) × 2 + (160 + 32) + (148 + 12) × 2 = 856px。 */''',
'''.option-zone{\n    width:148px;\n    margin-left:0;\n    margin-right:0;\n}\n\n/* 左側は CA名後から8px右へ寄せつつ、右端位置は維持。 */''',
'css option')
css=replace_once(css,
'''.attendance-zone{\n    float:left;\n    width:148px;\n    margin-right:12px;\n}''',
'''.attendance-zone{\n    float:left;\n    width:148px;\n    margin-left:12px;\n    margin-right:12px;\n}''',
'css attendance')
css_path.write_text(css,encoding='utf-8-sig')

# JS v28.23 -> v28.27
js=js_path.read_text(encoding='utf-8-sig')
js=replace_once(js,
'''var hasPositionedWindow = false;\nvar lastAvailWidth = 0;\nvar lastAvailHeight = 0;''',
'''var hasPositionedWindow = false;\nvar lastAvailLeft = 0;\nvar lastAvailTop = 0;\nvar lastAvailWidth = 0;\nvar lastAvailHeight = 0;''',
'js globals')
js=replace_once(js,
'''    lastAvailWidth=screen.availWidth;\n    lastAvailHeight=screen.availHeight;''',
'''    var initialWorkArea=getCurrentWorkArea();\n    lastAvailLeft=initialWorkArea.left;\n    lastAvailTop=initialWorkArea.top;\n    lastAvailWidth=initialWorkArea.width;\n    lastAvailHeight=initialWorkArea.height;''',
'js init area')
js=replace_once(js,
'''        try{\n            var w=screen.availWidth;\n            var h=screen.availHeight;\n\n            if(w!==lastAvailWidth || h!==lastAvailHeight){\n                lastAvailWidth=w;\n                lastAvailHeight=h;\n\n                // 解像度やWindows表示倍率の変更後に自動再フィット\n                resizeApp(true);\n            }''',
'''        try{\n            var area=getCurrentWorkArea();\n\n            if(area.left!==lastAvailLeft ||\n               area.top!==lastAvailTop ||\n               area.width!==lastAvailWidth ||\n               area.height!==lastAvailHeight){\n                lastAvailLeft=area.left;\n                lastAvailTop=area.top;\n                lastAvailWidth=area.width;\n                lastAvailHeight=area.height;\n\n                // 別モニターへ移動した場合や、解像度・表示倍率・タスクバー領域が変わった場合も\n                // 現在いるモニターの作業領域内へ再フィットする。\n                resizeApp(false);\n            }''',
'js screen watcher')
marker='''function centerCurrentWindow(outerW,outerH){\n    try{\n        var moveX=Math.max(0,Math.floor((screen.availWidth-outerW)/2));\n        var moveY=Math.max(0,Math.floor((screen.availHeight-outerH)/2));\n        window.moveTo(moveX,moveY);\n        hasPositionedWindow=true;\n    }catch(err){\n    }\n}\n'''
replacement='''function isFiniteScreenNumber(value){\n    var n=Number(value);\n    return !isNaN(n) && isFinite(n);\n}\n\nfunction getCurrentWindowPosition(){\n    var left=0;\n    var top=0;\n\n    try{\n        if(isFiniteScreenNumber(window.screenLeft)){\n            left=Number(window.screenLeft);\n        }else if(isFiniteScreenNumber(window.screenX)){\n            left=Number(window.screenX);\n        }\n    }catch(errLeft){\n    }\n\n    try{\n        if(isFiniteScreenNumber(window.screenTop)){\n            top=Number(window.screenTop);\n        }else if(isFiniteScreenNumber(window.screenY)){\n            top=Number(window.screenY);\n        }\n    }catch(errTop){\n    }\n\n    return {left:left,top:top};\n}\n\nfunction getCurrentWorkArea(){\n    var width=0;\n    var height=0;\n    var left=0;\n    var top=0;\n\n    try{\n        width=isFiniteScreenNumber(screen.availWidth) ? Number(screen.availWidth) : Number(screen.width||0);\n        height=isFiniteScreenNumber(screen.availHeight) ? Number(screen.availHeight) : Number(screen.height||0);\n\n        /*\n          v28.26:\n          availLeft / availTop はWindowsの仮想スクリーン座標上で、\n          現在のモニターの「タスクバー等を除いた作業領域」の左上を返す。\n          IE/HTA環境で未提供の場合は screen.left / screen.top をfallbackにする。\n        */\n        if(isFiniteScreenNumber(screen.availLeft)){\n            left=Number(screen.availLeft);\n        }else if(isFiniteScreenNumber(screen.left)){\n            left=Number(screen.left);\n        }\n\n        if(isFiniteScreenNumber(screen.availTop)){\n            top=Number(screen.availTop);\n        }else if(isFiniteScreenNumber(screen.top)){\n            top=Number(screen.top);\n        }\n    }catch(err){\n    }\n\n    if(width<=0){ width=1920; }\n    if(height<=0){ height=1080; }\n\n    return {left:left,top:top,width:width,height:height};\n}\n\nfunction fitCurrentWindowIntoWorkArea(fallbackW,fallbackH,preferUpperCenter){\n    try{\n        var area=getCurrentWorkArea();\n        var pos=getCurrentWindowPosition();\n        var outer=getWindowOuterSize(fallbackW,fallbackH);\n        var outerW=outer.width||fallbackW||0;\n        var outerH=outer.height||fallbackH||0;\n        var safeMargin=8;\n        var minLeft=area.left+safeMargin;\n        var minTop=area.top+safeMargin;\n        var maxLeft=area.left+area.width-outerW-safeMargin;\n        var maxTop=area.top+area.height-outerH-safeMargin;\n        var moveX=pos.left;\n        var moveY=pos.top;\n\n        /*\n          ウィンドウが作業領域より大きい場合は、まず作業領域内に収まるサイズへ補正。\n          通常はresizeApp側ですでに上限設定済みだが、DPI差などによる実寸差の保険。\n        */\n        if(outerW>area.width-(safeMargin*2) || outerH>area.height-(safeMargin*2)){\n            var correctedW=Math.min(outerW,Math.max(320,area.width-(safeMargin*2)));\n            var correctedH=Math.min(outerH,Math.max(240,area.height-(safeMargin*2)));\n            window.resizeTo(correctedW,correctedH);\n            outerW=correctedW;\n            outerH=correctedH;\n            maxLeft=area.left+area.width-outerW-safeMargin;\n            maxTop=area.top+area.height-outerH-safeMargin;\n        }\n\n        if(maxLeft<minLeft){\n            minLeft=area.left;\n            maxLeft=area.left;\n        }\n        if(maxTop<minTop){\n            minTop=area.top;\n            maxTop=area.top;\n        }\n\n        if(moveX<minLeft){ moveX=minLeft; }\n        if(moveX>maxLeft){ moveX=maxLeft; }\n        if(moveY<minTop){ moveY=minTop; }\n        if(moveY>maxTop){ moveY=maxTop; }\n\n        /*\n          v28.27:\n          2件目/3件目の展開などでフォームが縦に伸びたときは、\n          単に下端へ収めるだけではなく「中央より少し上」を目安に持ち上げる。\n          すでにそれより上に置かれている場合は、ユーザーの位置を尊重して動かさない。\n        */\n        if(preferUpperCenter){\n            var freeHeight=Math.max(0,area.height-outerH-(safeMargin*2));\n            var preferredTop=area.top+safeMargin+Math.round(freeHeight*0.38);\n\n            if(preferredTop<minTop){ preferredTop=minTop; }\n            if(preferredTop>maxTop){ preferredTop=maxTop; }\n            if(moveY>preferredTop){ moveY=preferredTop; }\n        }\n\n        if(Math.abs(moveX-pos.left)>1 || Math.abs(moveY-pos.top)>1){\n            window.moveTo(Math.round(moveX),Math.round(moveY));\n        }\n    }catch(err){\n    }\n}\n\nfunction centerCurrentWindow(outerW,outerH){\n    try{\n        var area=getCurrentWorkArea();\n        var moveX=area.left+Math.max(0,Math.floor((area.width-outerW)/2));\n        var moveY=area.top+Math.max(0,Math.floor((area.height-outerH)/2));\n        window.moveTo(moveX,moveY);\n        hasPositionedWindow=true;\n        fitCurrentWindowIntoWorkArea(outerW,outerH);\n    }catch(err){\n    }\n}\n'''
js=replace_once(js,marker,replacement,'js work area funcs')
js=replace_once(js,
'''            // 24インチ 1920×1080 を想定し、画面端へ少し余白を残す\n            var maxOuterW=screen.availWidth-40;\n            var maxOuterH=screen.availHeight-34;''',
'''            // 現在いるモニターの作業領域（タスクバー等を除く）を基準に上限を決める。\n            var workArea=getCurrentWorkArea();\n            var maxOuterW=workArea.width-40;\n            var maxOuterH=workArea.height-34;''',
'js max area')
js=replace_once(js,
'''            // v28.23: 初回配置時だけ中央寄せする。\n            // 2件目/3件目の追加・取消・送信後リセット・画面監視による再計算では\n            // moveToを呼ばず、現在いるモニター上の位置を維持する。\n            if(centerOnFirst && !hasPositionedWindow){\n                centerCurrentWindow(wantedW,wantedH);\n            }\n            window.scrollTo(0,0);''',
'''            // v28.26:\n            // 初回は「現在いるモニター」の中央へ配置。\n            // 2件目/3件目の展開・取消・送信後リセット時は現在位置を尊重しつつ、\n            // 下端/右端などが作業領域からはみ出す分だけ自動で戻す。\n            if(centerOnFirst && !hasPositionedWindow){\n                centerCurrentWindow(wantedW,wantedH);\n            }else{\n                fitCurrentWindowIntoWorkArea(wantedW,wantedH,true);\n            }\n\n            // resizeTo直後はHTAフレームの実寸反映が1テンポ遅れる場合があるため、\n            // 非表示中にもう一度実寸で補正してから表示する。\n            window.setTimeout(function(){\n                fitCurrentWindowIntoWorkArea(wantedW,wantedH,!centerOnFirst || hasPositionedWindow);\n            },10);\n\n            window.scrollTo(0,0);''',
'js resize positioning')
js_path.write_text(js,encoding='utf-8-sig')

# README: clearly mark v28.27 as repository latest without touching runtime behavior.
readme=readme_path.read_text(encoding='utf-8-sig')
header='''メンション依頼フォーム HTA 本番版 v28.27\n============================================================\n\n■ v28.27 最新版\n\n・依頼2／3展開時など、フォームが縦に伸びた際は現在のモニター内へ自動補正\n・補正位置は中央より少し上を目安にし、上下に余白が見える配置を優先\n・すでに十分上側に置かれている場合は、その位置を尊重\n・v28.25までの中央レイアウト調整（出社状況／代理CA名／オプションの間隔）を継承\n・CSV、送信、Pending、Popup等の機能仕様はv28.23系から維持\n\n'''
if not readme.startswith('メンション依頼フォーム HTA 本番版 v28.27'):
    readme=header+'\n'+readme
readme_path.write_text(readme,encoding='utf-8-sig')

# verify runtime files exactly match approved v28.27 package SHAs
expected={
    css_path:'9f4b0a471d2850feb1781e2acaee60a519dc25cd',
    js_path:'a9e079ab4bfbad00e1c74e004205a4276ca16763',
}
for p,sha in expected.items():
    b=p.read_bytes()
    actual=hashlib.sha1(b'blob '+str(len(b)).encode()+b'\0'+b).hexdigest()
    if actual!=sha:
        raise SystemExit(f'SHA_MISMATCH {p}: {actual} != {sha}')
print('PATCH_OK')
