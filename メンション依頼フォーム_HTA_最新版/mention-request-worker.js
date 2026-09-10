(function () {
    var FOR_READING = 1;
    var FOR_WRITING = 2;
    var FOR_APPENDING = 8;
    var TRISTATE_USE_DEFAULT = -2;

    var RETRY_MIN_MS = 450;
    var RETRY_MAX_MS = 1250;
    var MAX_RETRY_MS = 300000;
    var STALE_LOCK_MS = 600000;
    var COMPLETION_READY_SUFFIX = ".notify-ready";
    var COMPLETION_READY_MAX_MS = 30000;
    var COMPLETION_READY_POLL_MS = 200;
    var FAILURE_READY_RECENT_MS = 60000;

    var HEADER_LATEST = [
        "\u9001\u4fe1\u65e5\u6642",
        "RequestID",
        "\u62e0\u70b9",
        "\u4f9d\u983c\u8005",
        "\u4f9d\u983c\u756a\u53f7",
        "\u7d44\u7e54",
        "CA\u540d",
        "\u51fa\u793e\u72b6\u6cc1",
        "\u4ee3\u7406CA1",
        "\u4ee3\u7406CA2",
        "\u4ee3\u7406CA3",
        "\u30aa\u30d7\u30b7\u30e7\u30f3",
        "\u30e1\u30fc\u30eb\u30e1\u30e2",
        "\u51e6\u7406\u65e5\u6642",
        "\u671f\u65e5\u307e\u305f\u306f\u9762\u63a5\u65e5\u7a0b",
        "\u30bf\u30a4\u30d7"
    ];

    var HEADER_LEGACY = [
        "\u9001\u4fe1\u65e5\u6642",
        "RequestID",
        "\u62e0\u70b9",
        "\u4f9d\u983c\u8005",
        "\u4f9d\u983c\u756a\u53f7",
        "\u7d44\u7e54",
        "CA\u540d",
        "\u4ee3\u7406CA\u7d44\u7e54",
        "\u4ee3\u7406CA1",
        "\u4ee3\u7406CA2",
        "\u4ee3\u7406CA3",
        "\u30e1\u30fc\u30eb\u30e1\u30e2",
        "\u51e6\u7406\u65e5\u6642",
        "\u671f\u65e5",
        "\u6642\u77ed",
        "\u81f3\u6025",
        "\u30bf\u30a4\u30d7"
    ];

    var args = WScript.Arguments;
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var shell = new ActiveXObject("WScript.Shell");

    if (args.length < 2) {
        WScript.Quit(2);
    }

    var pendingPath = String(args.Item(0) || "");
    var csvFolder = String(args.Item(1) || "");
    var notBeforeMs = 0;

    if (args.length >= 3) {
        notBeforeMs = parseInt(String(args.Item(2) || "0"), 10);
        if (isNaN(notBeforeMs) || notBeforeMs < 0) {
            notBeforeMs = 0;
        }
    }

    if (!pendingPath || !fso.FileExists(pendingPath)) {
        WScript.Quit(0);
    }

    if (!csvFolder || !fso.FolderExists(csvFolder)) {
        showFailurePopupWhenReady(pendingPath);
        WScript.Quit(3);
    }

    // v28.5: start shared-CSV processing immediately.
    // The third argument is still accepted for compatibility, but no LT delay is applied.

    var pending;
    try {
        pending = readPendingCsv(pendingPath);
    } catch (e0) {
        showFailurePopupWhenReady(pendingPath);
        WScript.Quit(4);
    }

    var targetPath = "";
    try {
        targetPath = buildTargetPath(csvFolder, pending.baseName, pending.sentAt, pending.schema);
    } catch (eTarget) {
        showFailurePopupWhenReady(pendingPath);
        WScript.Quit(7);
    }
    var lockPath = targetPath + ".jlock";
    var startedAt = new Date().getTime();

    while ((new Date().getTime() - startedAt) < MAX_RETRY_MS) {
        if (!fso.FileExists(pendingPath)) {
            WScript.Quit(0);
        }

        if (!tryAcquireLock(lockPath)) {
            sleepRandom();
            continue;
        }

        try {
            // Another worker may have completed this Pending while we waited.
            if (!fso.FileExists(pendingPath)) {
                releaseLock(lockPath);
                WScript.Quit(0);
            }

            writeMissingRowsAndVerify(targetPath, pending);

            // Keep the shared lock until Pending deletion completes.
            // This prevents duplicate success popups from concurrent workers.
            if (deletePendingWithRetry(pendingPath)) {
                releaseLock(lockPath);
                showSuccessPopupWhenReady(pendingPath);
                WScript.Quit(0);
            }

            releaseLock(lockPath);
            showFailurePopupWhenReady(pendingPath);
            WScript.Quit(5);
        } catch (e1) {
            releaseLock(lockPath);

            if ((new Date().getTime() - startedAt) >= MAX_RETRY_MS) {
                break;
            }

            sleepRandom();
        }
    }

    showFailurePopupWhenReady(pendingPath);
    WScript.Quit(6);


    function readPendingCsv(path) {
        var text = readText(path);
        var rows = parseCsv(text);

        if (rows.length < 2) {
            throw new Error("PENDING_EMPTY");
        }

        var schema = detectHeader(rows[0]);
        var expectedColumns = schema === "latest" ? HEADER_LATEST.length : HEADER_LEGACY.length;

        var records = [];
        var requestId = "";
        var baseName = "";
        var sentAt = "";
        var requestNos = {};
        var i, row, no;

        for (i = 1; i < rows.length; i++) {
            row = rows[i];

            if (isEmptyRecord(row)) {
                continue;
            }

            if (row.length !== expectedColumns) {
                throw new Error("PENDING_BAD_COLUMNS");
            }

            if (!requestId) {
                sentAt = String(row[0] || "");
                requestId = String(row[1] || "");
                baseName = String(row[2] || "");

                if (!sentAt || !requestId || !baseName) {
                    throw new Error("PENDING_KEY_MISSING");
                }
            } else {
                if (String(row[1] || "") !== requestId ||
                    String(row[2] || "") !== baseName) {
                    throw new Error("PENDING_MIXED_PACKAGE");
                }
            }

            no = String(row[4] || "");
            if (!no) {
                throw new Error("REQUEST_NO_MISSING");
            }

            if (requestNos[no]) {
                throw new Error("PENDING_DUPLICATE_REQUEST_NO");
            }

            requestNos[no] = true;
            records.push(row);
        }

        if (records.length < 1) {
            throw new Error("PENDING_NO_DATA");
        }

        return {
            sentAt: sentAt,
            requestId: requestId,
            baseName: baseName,
            schema: schema,
            records: records
        };
    }


    function writeMissingRowsAndVerify(targetPath, pending) {
        var existingText = "";
        var existingRows = [];
        var existingKeys = {};
        var targetExists = fso.FileExists(targetPath);
        var targetHasContent = false;
        var i, row, key;

        if (targetExists) {
            existingText = readText(targetPath);
            targetHasContent = existingText.length > 0;

            if (targetHasContent) {
                existingRows = parseCsv(existingText);

                if (existingRows.length < 1) {
                    throw new Error("TARGET_PARSE_FAILED");
                }

                validateHeader(existingRows[0], pending.schema);

                for (i = 1; i < existingRows.length; i++) {
                    row = existingRows[i];

                    if (isEmptyRecord(row)) {
                        continue;
                    }

                    if (row.length !== expectedColumnCount(pending.schema)) {
                        throw new Error("TARGET_BAD_COLUMNS");
                    }

                    key = makeKey(row[1], row[4]);
                    existingKeys[key] = true;
                }
            }
        }

        var missing = [];

        for (i = 0; i < pending.records.length; i++) {
            row = pending.records[i];
            key = makeKey(row[1], row[4]);

            if (!existingKeys[key]) {
                missing.push(row);
            }
        }

        if (missing.length > 0) {
            appendRows(targetPath, existingText, targetHasContent, missing, pending.schema);
        }

        verifyAllRows(targetPath, pending);
    }


    function appendRows(targetPath, existingText, targetHasContent, rows, schema) {
        var stream, i;

        if (!targetHasContent) {
            stream = fso.OpenTextFile(
                targetPath,
                FOR_WRITING,
                true,
                TRISTATE_USE_DEFAULT
            );

            try {
                stream.WriteLine(headerLine(schema));

                for (i = 0; i < rows.length; i++) {
                    stream.WriteLine(csvLine(rows[i]));
                }
            } finally {
                stream.Close();
            }

            return;
        }

        stream = fso.OpenTextFile(
            targetPath,
            FOR_APPENDING,
            false,
            TRISTATE_USE_DEFAULT
        );

        try {
            if (!endsWithNewline(existingText)) {
                stream.WriteLine("");
            }

            for (i = 0; i < rows.length; i++) {
                stream.WriteLine(csvLine(rows[i]));
            }
        } finally {
            stream.Close();
        }
    }


    function verifyAllRows(targetPath, pending) {
        if (!fso.FileExists(targetPath)) {
            throw new Error("VERIFY_TARGET_MISSING");
        }

        var rows = parseCsv(readText(targetPath));
        var keys = {};
        var i, row, key;

        if (rows.length < 1) {
            throw new Error("VERIFY_EMPTY");
        }

        validateHeader(rows[0], pending.schema);

        for (i = 1; i < rows.length; i++) {
            row = rows[i];

            if (isEmptyRecord(row)) {
                continue;
            }

            if (row.length !== expectedColumnCount(pending.schema)) {
                throw new Error("VERIFY_BAD_COLUMNS");
            }

            keys[makeKey(row[1], row[4])] = true;
        }

        for (i = 0; i < pending.records.length; i++) {
            row = pending.records[i];
            key = makeKey(row[1], row[4]);

            if (!keys[key]) {
                throw new Error("VERIFY_FAILED");
            }
        }
    }


    function tryAcquireLock(path) {
        clearStaleLock(path);

        try {
            fso.CreateFolder(path);
            return true;
        } catch (e) {
            return false;
        }
    }


    function releaseLock(path) {
        try {
            if (fso.FolderExists(path)) {
                fso.DeleteFolder(path, true);
            }
        } catch (e) {
        }
    }


    function clearStaleLock(path) {
        try {
            if (!fso.FolderExists(path)) {
                return;
            }

            var folder = fso.GetFolder(path);
            var created = new Date(folder.DateCreated);
            var age = new Date().getTime() - created.getTime();

            if (age > STALE_LOCK_MS) {
                fso.DeleteFolder(path, true);
            }
        } catch (e) {
        }
    }


    function deletePendingWithRetry(path) {
        var i;

        if (!fso.FileExists(path)) {
            return true;
        }

        for (i = 0; i < 5; i++) {
            try {
                fso.DeleteFile(path, true);
                return true;
            } catch (e) {
                WScript.Sleep(250 + Math.floor(Math.random() * 500));
            }
        }

        return !fso.FileExists(path);
    }


    function showSuccessPopupWhenReady(pendingPath) {
        var readyPath = String(pendingPath || "") + COMPLETION_READY_SUFFIX;
        var startedAt = new Date().getTime();

        while (!fso.FileExists(readyPath) &&
               (new Date().getTime() - startedAt) < COMPLETION_READY_MAX_MS) {
            WScript.Sleep(COMPLETION_READY_POLL_MS);
        }

        showSuccessPopup();

        try {
            if (fso.FileExists(readyPath)) {
                fso.DeleteFile(readyPath, true);
            }
        } catch (e) {
        }
    }


    function showSuccessPopup() {
        try {
            var message =
                "\u30e1\u30f3\u30b7\u30e7\u30f3\u4f9d\u983c\u306e\u9001\u4fe1\u304c\u5b8c\u4e86\u3057\u307e\u3057\u305f\u3002" +
                "\r\n\r\n" +
                "\u5171\u6709CSV\u3078\u306e\u66f8\u304d\u8fbc\u307f\u3068\u78ba\u8a8d\u304c\u5b8c\u4e86\u3057\u3066\u3044\u307e\u3059\u3002";

            // Timeout 0: keep the completion notice visible until OK is pressed.
            // 64      = MB_ICONINFORMATION
            // 65536   = MB_SETFOREGROUND
            // 262144  = MB_TOPMOST
            // This brings the completion notice to the front even when the main HTA is minimized.
            shell.Popup(
                message,
                0,
                "\u30e1\u30f3\u30b7\u30e7\u30f3\u4f9d\u983c\u30d5\u30a9\u30fc\u30e0 | \u9001\u4fe1\u5b8c\u4e86",
                327744
            );
        } catch (e) {
            // Notification failure must never affect CSV delivery.
        }
    }


    function shouldWaitForFailureReady(pendingPath) {
        try {
            if (!pendingPath || !fso.FileExists(pendingPath)) {
                return false;
            }

            var file = fso.GetFile(pendingPath);
            var createdAt = new Date(file.DateCreated).getTime();
            var age = new Date().getTime() - createdAt;

            return age >= 0 && age <= FAILURE_READY_RECENT_MS;
        } catch (e) {
            return false;
        }
    }


    function showFailurePopupWhenReady(pendingPath) {
        var readyPath = String(pendingPath || "") + COMPLETION_READY_SUFFIX;
        var startedAt = new Date().getTime();

        // For a just-submitted request, avoid overlapping the acceptance modal.
        // Startup/manual retries are normally old Pending files, so notify immediately.
        if (shouldWaitForFailureReady(pendingPath)) {
            while (!fso.FileExists(readyPath) &&
                   (new Date().getTime() - startedAt) < COMPLETION_READY_MAX_MS) {
                WScript.Sleep(COMPLETION_READY_POLL_MS);
            }
        }

        showFailurePopup();

        try {
            if (fso.FileExists(readyPath)) {
                fso.DeleteFile(readyPath, true);
            }
        } catch (e) {
        }
    }


    function showFailurePopup() {
        try {
            var message =
                "\u30e1\u30f3\u30b7\u30e7\u30f3\u4f9d\u983c\u306e\u9001\u4fe1\u3092\u5b8c\u4e86\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002" +
                "\r\n\r\n" +
                "\u672a\u9001\u4fe1\u30c7\u30fc\u30bf\u306f\u4fdd\u5b58\u3055\u308c\u3066\u3044\u307e\u3059\u3002" +
                "\r\n" +
                "\u30d5\u30a9\u30fc\u30e0\u3092\u958b\u304d\u3001\u300c\u672a\u9001\u4fe1\u3092\u518d\u9001\u300d\u304b\u3089\u518d\u9001\u3057\u3066\u304f\u3060\u3055\u3044\u3002";

            // 16      = MB_ICONERROR
            // 65536   = MB_SETFOREGROUND
            // 262144  = MB_TOPMOST
            // Timeout 0 keeps the failure notice visible until OK is pressed.
            shell.Popup(
                message,
                0,
                "\u30e1\u30f3\u30b7\u30e7\u30f3\u4f9d\u983c\u30d5\u30a9\u30fc\u30e0 | \u9001\u4fe1\u5931\u6557",
                327696
            );
        } catch (e) {
            // Notification failure must never alter Pending/CSV state.
        }
    }


    function buildTargetPath(folder, baseName, sentAt, schema) {
        var dateKey = dateKeyFromSentAt(sentAt);
        var safeBase = safeFileName(baseName);
        var preferred = fso.BuildPath(folder, safeBase + "_" + dateKey + ".csv");

        if (!fso.FileExists(preferred) || readText(preferred).length === 0) {
            return preferred;
        }

        var currentSchema = detectHeader(parseCsv(readText(preferred))[0]);
        if (currentSchema === schema) {
            return preferred;
        }

        // Rollout-day safety: never mix the old and latest schemas in one CSV.
        // Existing legacy daily files stay untouched; latest rows use a suffixed file.
        if (schema === "latest") {
            return fso.BuildPath(folder, safeBase + "_" + dateKey + "_v28.22.csv");
        }

        return fso.BuildPath(folder, safeBase + "_" + dateKey + "_legacy.csv");
    }


    function dateKeyFromSentAt(value) {
        var m = /^(\d{4})\/(\d{2})\/(\d{2})/.exec(String(value || ""));

        if (!m) {
            throw new Error("DATE_INVALID");
        }

        return m[1] + m[2] + m[3];
    }


    function safeFileName(value) {
        return String(value || "").replace(/[\\\/:\*\?"<>\|]/g, "_");
    }


    function makeKey(requestId, requestNo) {
        return String(requestId || "") + "|" + String(requestNo || "");
    }


    function readText(path) {
        var stream = fso.OpenTextFile(
            path,
            FOR_READING,
            false,
            TRISTATE_USE_DEFAULT
        );

        try {
            return String(stream.ReadAll() || "");
        } finally {
            stream.Close();
        }
    }


    function normalizeHeaderField(value, index) {
        var actual = String(value || "");
        if (index === 0 && actual.length > 0 && actual.charCodeAt(0) === 0xFEFF) {
            actual = actual.substring(1);
        }
        return actual;
    }


    function headerMatches(row, header) {
        var i;
        if (!row || row.length !== header.length) {
            return false;
        }
        for (i = 0; i < header.length; i++) {
            if (normalizeHeaderField(row[i], i) !== header[i]) {
                return false;
            }
        }
        return true;
    }


    function detectHeader(row) {
        if (headerMatches(row, HEADER_LATEST)) { return "latest"; }
        if (headerMatches(row, HEADER_LEGACY)) { return "legacy"; }
        throw new Error("HEADER_MISMATCH");
    }


    function validateHeader(row, schema) {
        var detected = detectHeader(row);
        if (detected !== schema) {
            throw new Error("HEADER_SCHEMA_MISMATCH");
        }
    }


    function expectedColumnCount(schema) {
        return schema === "latest" ? HEADER_LATEST.length : HEADER_LEGACY.length;
    }


    function headerLine(schema) {
        return (schema === "latest" ? HEADER_LATEST : HEADER_LEGACY).join(",");
    }


    function csvLine(fields) {
        var out = [];
        var i;

        for (i = 0; i < fields.length; i++) {
            out.push(csvEscape(fields[i]));
        }

        return out.join(",");
    }


    function csvEscape(value) {
        var s = String(value == null ? "" : value);
        return '"' + s.replace(/"/g, '""') + '"';
    }


    function parseCsv(text) {
        text = String(text == null ? "" : text);

        var rows = [];
        var row = [];
        var field = "";
        var inQuotes = false;
        var i, ch, next;

        for (i = 0; i < text.length; i++) {
            ch = text.charAt(i);

            if (inQuotes) {
                if (ch === '"') {
                    next = (i + 1 < text.length) ? text.charAt(i + 1) : "";

                    if (next === '"') {
                        field += '"';
                        i++;
                    } else {
                        inQuotes = false;
                    }
                } else {
                    field += ch;
                }

                continue;
            }

            if (ch === '"') {
                inQuotes = true;
            } else if (ch === ",") {
                row.push(field);
                field = "";
            } else if (ch === "\r" || ch === "\n") {
                row.push(field);
                field = "";
                rows.push(row);
                row = [];

                if (ch === "\r" &&
                    i + 1 < text.length &&
                    text.charAt(i + 1) === "\n") {
                    i++;
                }
            } else {
                field += ch;
            }
        }

        if (inQuotes) {
            throw new Error("CSV_UNCLOSED_QUOTE");
        }

        if (field !== "" || row.length > 0) {
            row.push(field);
            rows.push(row);
        }

        return rows;
    }


    function isEmptyRecord(row) {
        var i;

        if (!row || row.length === 0) {
            return true;
        }

        for (i = 0; i < row.length; i++) {
            if (String(row[i] || "") !== "") {
                return false;
            }
        }

        return true;
    }


    function endsWithNewline(text) {
        if (!text) {
            return false;
        }

        var ch = text.charAt(text.length - 1);
        return ch === "\r" || ch === "\n";
    }


    function sleepRandom() {
        var span = RETRY_MAX_MS - RETRY_MIN_MS + 1;
        var ms = RETRY_MIN_MS + Math.floor(Math.random() * span);
        WScript.Sleep(ms);
    }
})();