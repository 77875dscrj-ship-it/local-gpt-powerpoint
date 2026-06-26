(function () {
  var base = "https://localhost:8765";
  var statusEl = document.getElementById("status");
  var contextEl = document.getElementById("pptContext");
  var chatEl = document.getElementById("chat");
  var promptEl = document.getElementById("prompt");
  var composerEl = document.getElementById("composer");
  var sourceFileEl = document.getElementById("sourceFile");
  var clipboardImageButtonEl = document.getElementById("clipboardImageButton");
  var attachmentListEl = document.getElementById("attachmentList");
  var permissionCardEl = document.getElementById("permissionCard");
  var planSummaryEl = document.getElementById("planSummary");
  var stepListEl = document.getElementById("stepList");
  var outlineListEl = document.getElementById("outlineList");
  var previewPanelEl = document.getElementById("previewPanel");
  var previewListEl = document.getElementById("previewList");
  var commitPreviewEl = document.getElementById("commitPreview");
  var rollbackLastEl = document.getElementById("rollbackLast");

  var history = [];
  var attachments = [];
  var pastedImageCounter = 0;
  var lastImagePasteAt = 0;
  var lastClipboardShortcutAt = 0;
  var clipboardFallbackBusy = false;
  var pendingPlan = null;
  var pendingTransactionId = null;
  var lastTransactionId = null;
  var busy = false;

  function setStatus(text) {
    statusEl.innerHTML = escapeHtml(text);
  }

  function resizeChat() {
    var viewport = window.innerHeight || document.documentElement.clientHeight || document.body.clientHeight || 0;
    if (!viewport) return;
    var composerHeight = composerEl ? composerEl.offsetHeight : 0;
    var permissionHeight = permissionCardEl && permissionCardEl.className.indexOf("hidden") < 0 ? permissionCardEl.offsetHeight : 0;
    var top = chatEl.offsetTop || 0;
    var available = viewport - top - composerHeight - permissionHeight - 18;
    if (available < 220) available = 220;
    chatEl.style.height = available + "px";
  }

  function escapeHtml(text) {
    return String(text || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function xhr(method, url, data, cb) {
    var req = new XMLHttpRequest();
    req.open(method, base + url, true);
    req.setRequestHeader("Content-Type", "application/json;charset=UTF-8");
    req.onreadystatechange = function () {
      if (req.readyState !== 4) return;
      var payload = null;
      try {
        payload = req.responseText ? JSON.parse(req.responseText) : {};
      } catch (e) {
        cb(new Error("응답을 읽지 못했습니다."));
        return;
      }
      if (req.status < 200 || req.status >= 300 || payload.ok === false) {
        cb(new Error(payload.error || ("HTTP " + req.status)));
        return;
      }
      cb(null, payload);
    };
    req.onerror = function () {
      cb(new Error("로컬 서버에 연결하지 못했습니다."));
    };
    req.send(data ? JSON.stringify(data) : null);
  }

  function addMessage(role, text, extraClass) {
    var row = document.createElement("div");
    row.className = "message " + role + (extraClass ? " " + extraClass : "");
    var bubble = document.createElement("div");
    bubble.className = "bubble";
    bubble.innerHTML = escapeHtml(text).replace(/\n/g, "<br>");
    row.appendChild(bubble);
    chatEl.appendChild(row);
    chatEl.scrollTop = chatEl.scrollHeight;
    resizeChat();
    if (role === "user" || role === "assistant") {
      history.push({ role: role, content: text });
      if (history.length > 16) history = history.slice(history.length - 16);
    }
    return bubble;
  }

  function updateAttachmentList() {
    attachmentListEl.innerHTML = "";
    for (var i = 0; i < attachments.length; i++) {
      var chip = document.createElement("div");
      chip.className = "attachmentChip" + (attachments[i].kind === "image" ? " imageAttachment" : "");
      var remove = "<button class=\"attachmentRemove\" data-index=\"" + i + "\" title=\"제거\">x</button>";
      if (attachments[i].kind === "image") {
        chip.innerHTML = "<img src=\"" + escapeHtml(attachments[i].previewUrl || "") + "\" alt=\"\">" +
          "<span>" + escapeHtml(attachments[i].name) + "</span>" + remove;
      } else {
        chip.innerHTML = "<span>" + escapeHtml(attachments[i].name) + "</span>" + remove;
      }
      attachmentListEl.appendChild(chip);
    }
    resizeChat();
  }

  function sourceText() {
    var parts = [];
    for (var i = 0; i < attachments.length; i++) {
      if (attachments[i].kind === "image") continue;
      parts.push("[Source: " + attachments[i].name + "]\n" + attachments[i].text);
    }
    return parts.join("\n\n---\n\n");
  }

  function imageAttachments() {
    var images = [];
    for (var i = 0; i < attachments.length; i++) {
      if (attachments[i].kind === "image") {
        images.push({
          name: attachments[i].name,
          path: attachments[i].path,
          mimeType: attachments[i].mimeType,
          sizeBytes: attachments[i].sizeBytes
        });
      }
    }
    return images;
  }

  function refreshHealth() {
    xhr("GET", "/api/health", null, function (err, data) {
      if (err) {
        setStatus("서버 연결 실패");
        return;
      }
      var model = data.effectiveModel || (data.codex && data.codex.model) || "Codex 현재 설정";
      setStatus("Codex OAuth · " + model);
    });
  }

  function refreshContext() {
    xhr("GET", "/api/ppt/selection", null, function (err, data) {
      if (err) {
        contextEl.innerHTML = "PowerPoint 문맥 없음: " + escapeHtml(err.message);
        return;
      }
      var sel = data.selection || {};
      var text = "슬라이드 " + data.slideIndex + "/" + data.slideCount;
      if (data.slideId) text += " · ID " + data.slideId;
      if (sel.shapes && sel.shapes.length) {
        text += " · 선택 블럭 " + sel.shapes.length + "개";
        if (sel.shapes[0].fontSize) text += " · 글꼴 " + sel.shapes[0].fontSize + "pt";
      } else if (sel.type === 3) {
        text += " · 선택 텍스트";
      } else {
        text += " · 선택 없음";
      }
      contextEl.innerHTML = escapeHtml(text);
    });
  }

  function setBusy(value, text) {
    busy = value;
    document.getElementById("send").disabled = value;
    document.getElementById("allowEdit").disabled = value;
    commitPreviewEl.disabled = value;
    document.getElementById("attachButton").disabled = value;
    if (clipboardImageButtonEl) clipboardImageButtonEl.disabled = value;
    if (text) setStatus(text);
  }

  function requestedMode() {
    var radios = document.getElementsByName("requestedMode");
    for (var i = 0; i < radios.length; i++) {
      if (radios[i].checked) return radios[i].value;
    }
    return "review";
  }

  function riskLabel(risk) {
    if (risk === "high") return "높음";
    if (risk === "medium") return "중간";
    return "낮음";
  }

  function changeLocation(item, index) {
    var ref = item.slideRef || {};
    if (ref.tempSlideKey) return "새 슬라이드";
    if (ref.slideIndex) return "슬라이드 " + ref.slideIndex;
    if (ref.slideId) return "SlideID " + ref.slideId;
    return "변경 " + (index + 1);
  }

  function renderOutline(plan) {
    outlineListEl.innerHTML = "";
    var outline = plan.outline || [];
    if (!outline.length) {
      var empty = document.createElement("div");
      empty.className = "outlineEmpty";
      empty.innerHTML = "적용할 슬라이드 변경은 없습니다.";
      outlineListEl.appendChild(empty);
      return;
    }

    for (var i = 0; i < outline.length; i++) {
      var item = outline[i] || {};
      var card = document.createElement("div");
      card.className = "outlineCard risk-" + (item.risk || "low");

      var head = document.createElement("label");
      head.className = "outlineHead";

      var checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.className = "changeCheck";
      checkbox.setAttribute("data-change-id", item.changeId || "");
      checkbox.checked = item.selected !== false;
      head.appendChild(checkbox);

      var titleBox = document.createElement("span");
      titleBox.className = "outlineTitleBox";
      var title = item.title || "슬라이드 변경";
      titleBox.innerHTML =
        "<strong>" + escapeHtml(changeLocation(item, i)) + " · " + escapeHtml(title) + "</strong>" +
        "<em>위험도 " + escapeHtml(riskLabel(item.risk)) + "</em>";
      head.appendChild(titleBox);
      card.appendChild(head);

      if (item.keyMessage) {
        var key = document.createElement("div");
        key.className = "outlineKey";
        key.innerHTML = escapeHtml(item.keyMessage);
        card.appendChild(key);
      }

      var changes = item.changes || [];
      if (changes.length) {
        var list = document.createElement("ul");
        list.className = "changeList";
        for (var j = 0; j < changes.length; j++) {
          var li = document.createElement("li");
          li.innerHTML = escapeHtml(changes[j].summary || changes[j].type || "변경");
          list.appendChild(li);
        }
        card.appendChild(list);
      }

      if (item.rationale) {
        var why = document.createElement("div");
        why.className = "outlineMeta";
        why.innerHTML = "근거: " + escapeHtml(item.rationale);
        card.appendChild(why);
      }

      var refs = item.sourceRefs || [];
      if (refs.length) {
        var src = document.createElement("div");
        src.className = "outlineMeta";
        var labels = [];
        for (var k = 0; k < refs.length; k++) {
          labels.push((refs[k].sourceId || "source") + (refs[k].locator ? " / " + refs[k].locator : ""));
        }
        src.innerHTML = "출처: " + escapeHtml(labels.join(", "));
        card.appendChild(src);
      }

      outlineListEl.appendChild(card);
    }
  }

  function selectedChangeIds() {
    var checks = outlineListEl.getElementsByTagName("input");
    var ids = [];
    for (var i = 0; i < checks.length; i++) {
      if (checks[i].className.indexOf("changeCheck") >= 0 && checks[i].checked) {
        var id = checks[i].getAttribute("data-change-id");
        if (id) ids.push(id);
      }
    }
    return ids;
  }

  function setAllChanges(value) {
    var checks = outlineListEl.getElementsByTagName("input");
    for (var i = 0; i < checks.length; i++) {
      if (checks[i].className.indexOf("changeCheck") >= 0) checks[i].checked = value;
    }
  }

  function renderPermission(plan) {
    pendingPlan = plan;
    pendingTransactionId = null;
    previewPanelEl.className = "previewPanel hidden";
    previewListEl.innerHTML = "";
    commitPreviewEl.className = "hidden";
    document.getElementById("allowEdit").className = "";
    planSummaryEl.innerHTML = escapeHtml(plan.assistantMessage || "작업 계획을 만들었습니다.").replace(/\n/g, "<br>");
    stepListEl.innerHTML = "";
    var steps = plan.steps || [];
    for (var i = 0; i < steps.length; i++) {
      var li = document.createElement("li");
      var title = steps[i].title || ("단계 " + (i + 1));
      var detail = steps[i].detail || "";
      li.innerHTML = "<strong>" + escapeHtml(title) + "</strong>" + (detail ? "<span>" + escapeHtml(detail) + "</span>" : "");
      stepListEl.appendChild(li);
    }
    if (!steps.length) {
      var empty = document.createElement("li");
      empty.innerHTML = "<strong>검토</strong><span>적용할 PowerPoint 변경은 없습니다.</span>";
      stepListEl.appendChild(empty);
    }
    renderOutline(plan);
    if (plan.needsPermission || (plan.actions && plan.actions.length)) {
      permissionCardEl.className = "permission";
    } else {
      permissionCardEl.className = "permission hidden";
    }
    resizeChat();
  }

  function hidePermission() {
    pendingPlan = null;
    pendingTransactionId = null;
    permissionCardEl.className = "permission hidden";
    resizeChat();
  }

  function sendPrompt() {
    if (busy) return;
    var message = promptEl.value.replace(/^\s+|\s+$/g, "");
    var images = imageAttachments();
    if (!message && !images.length) return;
    if (!message) message = "첨부한 이미지를 참고해 주세요.";
    hidePermission();
    addMessage("user", message + (images.length ? "\n[이미지 첨부: " + images.length + "개]" : ""));
    promptEl.value = "";
    var thinking = addMessage("assistant", "생각 중입니다. 현재 덱과 선택 항목을 읽고 작업 단계를 짜고 있어요.", "thinking");
    setBusy(true, "작업 계획 작성 중");
    xhr("POST", "/api/chat/plan", {
      message: message,
      source: sourceText(),
      images: images,
      history: history,
      requestedMode: requestedMode()
    }, function (err, data) {
      setBusy(false, err ? "오류" : "계획 생성됨");
      if (thinking && thinking.parentNode) thinking.parentNode.parentNode.removeChild(thinking.parentNode);
      if (err) {
        addMessage("assistant", "오류: " + err.message);
        return;
      }
      var plan = data.plan || {};
      addMessage("assistant", plan.assistantMessage || "작업 계획을 만들었습니다.");
      renderPermission(plan);
    });
  }

  function renderPreview(data) {
    var slides = data.slides || [];
    pendingTransactionId = data.transactionId;
    previewListEl.innerHTML = "";
    for (var i = 0; i < slides.length; i++) {
      var slide = slides[i] || {};
      var card = document.createElement("div");
      card.className = "previewCard";
      var title = "슬라이드 " + (slide.slideIndex || "?");
      if (slide.slideId) title += " · ID " + slide.slideId;
      var html = "<div class=\"previewSlideTitle\">" + escapeHtml(title) + "</div>";
      if (slide.beforeImageUrl) {
        html += "<div class=\"previewLabel\">변경 전</div><img src=\"" + escapeHtml(base + slide.beforeImageUrl) + "\">";
      }
      if (slide.afterImageUrl) {
        html += "<div class=\"previewLabel\">변경 후</div><img src=\"" + escapeHtml(base + slide.afterImageUrl) + "\">";
      }
      card.innerHTML = html;
      previewListEl.appendChild(card);
    }
    if (!slides.length) {
      previewListEl.innerHTML = "<div class=\"outlineEmpty\">표시할 미리보기 이미지가 없습니다.</div>";
    }
    previewPanelEl.className = "previewPanel";
    document.getElementById("allowEdit").className = "hidden";
    commitPreviewEl.className = "";
    resizeChat();
  }

  function createPreview() {
    if (busy || !pendingPlan) return;
    var plan = pendingPlan;
    var ids = selectedChangeIds();
    if ((plan.actions && plan.actions.length) && ids.length < 1) {
      addMessage("assistant", "적용할 변경을 하나 이상 선택해 주세요.");
      return;
    }
    addMessage("assistant", "선택한 변경을 복사본에서 먼저 미리보기로 생성합니다.", "thinking");
    setBusy(true, "미리보기 생성 중");
    xhr("POST", "/api/plans/" + encodeURIComponent(plan.planId) + "/commit", { selectedChangeIds: ids }, function (err, data) {
      setBusy(false, err ? "오류" : "미리보기 완료");
      if (err) {
        addMessage("assistant", "오류: " + err.message);
        return;
      }
      if (data && data.applied) {
        handleCommitSuccess(data);
        return;
      }
      renderPreview(data);
      addMessage("assistant", "미리보기를 만들었습니다. 현재 덱은 아직 수정하지 않았습니다.");
    });
  }

  function appendApplyResultLines(lines, results) {
    for (var i = 0; i < results.length; i++) {
      var r = results[i];
      var line = "- " + (r.type || "slide");
      if (r.slide) line += " · slide " + r.slide;
      if (r.target) line += " · " + r.target;
      if (r.noOp || (r.type === "format_selection" && Number(r.changed || 0) === 0)) {
        line += " · 변경 없음";
        if (r.noOpReason) line += " (" + r.noOpReason + ")";
      } else if (r.changed) {
        line += " · " + r.changed + "개 변경";
      }
      if (r.changedProperties && r.changedProperties.length) line += " · " + r.changedProperties.join(", ");
      lines.push(line);
    }
  }

  function handleCommitSuccess(data) {
    lastTransactionId = data.transactionId || pendingTransactionId;
    var transaction = data.transaction || {};
    var backup = transaction.backup || {};
    rollbackLastEl.className = backup.created ? "miniButton" : "miniButton hidden";
    hidePermission();
    var result = data.result || {};
    var results = result.results || [];
    var target = transaction.commitTarget || null;
    var lines = result.noOp ? [result.noOpReason || "변경 없음: 적용할 새 서식 변경이 없습니다."] : ["완료됐습니다."];
    if (!backup.created) {
      lines.push("별도 저장본은 만들지 않았습니다. 되돌리려면 PowerPoint에서 Ctrl+Z를 사용하세요.");
    }
    if (target && target.mode === "editable_copy") {
      lines.push("원본은 수정하지 않았고, 편집용 복사본에 적용했습니다.");
      if (target.committedFullName || target.editableCopyPath) {
        lines.push("복사본: " + (target.committedFullName || target.editableCopyPath));
      }
    }
    appendApplyResultLines(lines, results);
    addMessage("assistant", lines.join("\n"));
    refreshContext();
  }

  function commitPreview() {
    if (busy || !pendingTransactionId) return;
    setBusy(true, "백업 후 적용 중");
    addMessage("assistant", "승인되었습니다. 원본 백업을 만든 뒤 PowerPoint에 적용합니다.", "thinking");
    xhr("POST", "/api/transactions/" + encodeURIComponent(pendingTransactionId) + "/commit", { approved: true }, function (err, data) {
      setBusy(false, err ? "오류" : "PowerPoint 적용됨");
      if (err) {
        addMessage("assistant", "오류: " + err.message);
        return;
      }
      lastTransactionId = pendingTransactionId;
      rollbackLastEl.className = "miniButton";
      hidePermission();
      var result = data.result || {};
      var results = result.results || [];
      var target = data.transaction && data.transaction.commitTarget ? data.transaction.commitTarget : null;
      var lines = result.noOp ? [result.noOpReason || "변경 없음: 적용할 새 서식 변경이 없습니다."] : ["완료했습니다."];
      if (target && target.mode === "editable_copy") {
        lines.push("원본은 수정하지 않았고, 편집용 복사본에 적용했습니다.");
        if (target.committedFullName || target.editableCopyPath) {
          lines.push("복사본: " + (target.committedFullName || target.editableCopyPath));
        }
      }
      appendApplyResultLines(lines, results);
      addMessage("assistant", lines.join("\n"));
      refreshContext();
    });
  }

  function rollbackLast() {
    if (busy || !lastTransactionId) return;
    setBusy(true, "백업본 여는 중");
    xhr("POST", "/api/transactions/" + encodeURIComponent(lastTransactionId) + "/rollback", {}, function (err, data) {
      setBusy(false, err ? "오류" : "백업본 열림");
      if (err) {
        addMessage("assistant", "오류: " + err.message);
        return;
      }
      addMessage("assistant", data.message || "백업본을 새 프레젠테이션으로 열었습니다.");
    });
  }

  function dataUrlBase64(dataUrl) {
    var value = String(dataUrl || "");
    var comma = value.indexOf(",");
    return comma >= 0 ? value.slice(comma + 1) : value;
  }

  function uploadImageFile(file, previewUrl, cb) {
    var name = file && file.name ? file.name : ("pasted-image-" + (++pastedImageCounter) + ".png");
    var mimeType = file && file.type ? file.type : "image/png";
    if (!/^image\//i.test(mimeType)) {
      cb(new Error("이미지 파일이 아닙니다."));
      return;
    }
    xhr("POST", "/api/images", {
      name: name,
      mimeType: mimeType,
      contentBase64: dataUrlBase64(previewUrl)
    }, function (err, data) {
      if (err) {
        cb(err);
        return;
      }
      attachments.push({
        kind: "image",
        name: data.name || name,
        mimeType: data.mimeType || mimeType,
        sizeBytes: data.sizeBytes || 0,
        path: data.path,
        previewUrl: previewUrl
      });
      lastImagePasteAt = (new Date()).getTime();
      updateAttachmentList();
      cb(null, data);
    });
  }

  function readAndUploadImage(file, cb) {
    var reader = new FileReader();
    reader.onload = function () {
      uploadImageFile(file, String(reader.result || ""), cb);
    };
    reader.onerror = function () {
      cb(new Error("이미지를 읽지 못했습니다."));
    };
    reader.readAsDataURL(file);
  }

  function addServerImageAttachment(data) {
    if (!data || !data.path) return false;
    attachments.push({
      kind: "image",
      name: data.name || "clipboard-image.png",
      mimeType: data.mimeType || "image/png",
      sizeBytes: data.sizeBytes || 0,
      path: data.path,
      previewUrl: data.url ? base + data.url : ""
    });
    lastImagePasteAt = (new Date()).getTime();
    updateAttachmentList();
    return true;
  }

  function importClipboardImageFromWindows(showNoImageMessage) {
    if (busy || clipboardFallbackBusy) return;
    clipboardFallbackBusy = true;
    if (clipboardImageButtonEl) clipboardImageButtonEl.disabled = true;
    setStatus("\uD074\uB9BD\uBCF4\uB4DC \uC774\uBBF8\uC9C0 \uD655\uC778 \uC911");
    xhr("POST", "/api/clipboard/image", {}, function (err, data) {
      clipboardFallbackBusy = false;
      if (clipboardImageButtonEl) clipboardImageButtonEl.disabled = busy;
      if (err) {
        if (showNoImageMessage) addMessage("assistant", "\uD074\uB9BD\uBCF4\uB4DC \uC774\uBBF8\uC9C0\uB97C \uC77D\uC9C0 \uBABB\uD588\uC2B5\uB2C8\uB2E4.\n" + err.message);
        setStatus("\uC624\uB958");
        return;
      }
      if (!data || !data.hasImage) {
        if (showNoImageMessage) addMessage("assistant", "\uD074\uB9BD\uBCF4\uB4DC\uC5D0 \uADF8\uB9BC\uC774 \uC5C6\uC2B5\uB2C8\uB2E4. PowerPoint\uC5D0\uC11C \uADF8\uB9BC\uC744 \uBCF5\uC0AC\uD55C \uB4A4 \uB2E4\uC2DC \uB20C\uB7EC\uC8FC\uC138\uC694.");
        setStatus("\uD074\uB9BD\uBCF4\uB4DC \uC774\uBBF8\uC9C0 \uC5C6\uC74C");
        return;
      }
      if (addServerImageAttachment(data)) {
        addMessage("assistant", "\uD074\uB9BD\uBCF4\uB4DC \uADF8\uB9BC\uC744 \uCCA8\uBD80\uD588\uC2B5\uB2C8\uB2E4. \uC774\uC81C \uC774 \uADF8\uB9BC\uC744 \uBCF4\uACE0 \uC791\uC5C5\uD558\uB77C\uACE0 \uC694\uCCAD\uD558\uBA74 \uB429\uB2C8\uB2E4.");
        setStatus("\uC774\uBBF8\uC9C0 \uCCA8\uBD80\uB428");
      }
    });
  }

  function scheduleClipboardFallbackCheck() {
    var now = (new Date()).getTime();
    if (now - lastClipboardShortcutAt < 500) return;
    lastClipboardShortcutAt = now;
    window.setTimeout(function () {
      var elapsed = (new Date()).getTime() - lastImagePasteAt;
      if (elapsed > 1200) importClipboardImageFromWindows(false);
    }, 250);
  }

  function handlePaste(e) {
    e = e || window.event;
    if (e.__localGptImagePasteHandled) return false;
    if (busy) return true;
    var data = e.clipboardData || window.clipboardData;
    if (!data) {
      importClipboardImageFromWindows(false);
      return true;
    }
    var items = data.items || [];
    var imageFiles = [];
    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      if (item && item.type && /^image\//i.test(item.type) && item.getAsFile) {
        var file = item.getAsFile();
        if (file) imageFiles.push(file);
      }
    }
    if (!imageFiles.length && data.files && data.files.length) {
      for (var j = 0; j < data.files.length; j++) {
        if (data.files[j].type && /^image\//i.test(data.files[j].type)) imageFiles.push(data.files[j]);
      }
    }
    if (!imageFiles.length) {
      importClipboardImageFromWindows(false);
      return true;
    }
    e.__localGptImagePasteHandled = true;
    if (e.preventDefault) e.preventDefault();
    setBusy(true, "이미지 붙여넣는 중");
    var index = 0;
    function next() {
      if (index >= imageFiles.length) {
        setBusy(false, "이미지 첨부됨");
        addMessage("assistant", "이미지를 첨부했습니다. 이제 이 이미지를 참고해서 요청할 수 있습니다.");
        return;
      }
      readAndUploadImage(imageFiles[index++], function (err) {
        if (err) {
          setBusy(false, "오류");
          addMessage("assistant", "이미지를 첨부하지 못했습니다: " + err.message);
          return;
        }
        next();
      });
    }
    next();
    return false;
  }

  function uploadFiles() {
    var files = sourceFileEl.files || [];
    if (!files.length || busy) return;
    var index = 0;
    function next() {
      if (index >= files.length) {
        sourceFileEl.value = "";
        updateAttachmentList();
        setStatus("소스 파일 추가됨");
        return;
      }
      var file = files[index++];
      setStatus("파일 읽는 중 " + index + "/" + files.length);
      if (file.type && /^image\//i.test(file.type)) {
        readAndUploadImage(file, function (err) {
          if (err) {
            addMessage("assistant", "이미지를 첨부하지 못했습니다: " + file.name + "\n" + err.message);
            setStatus("이미지 첨부 실패");
            return;
          }
          addMessage("assistant", "이미지를 첨부했습니다: " + file.name);
          next();
        });
        return;
      }
      var reader = new FileReader();
      reader.onload = function () {
        var dataUrl = String(reader.result || "");
        var comma = dataUrl.indexOf(",");
        var contentBase64 = comma >= 0 ? dataUrl.slice(comma + 1) : dataUrl;
        xhr("POST", "/api/source/extract", {
          name: file.name,
          contentBase64: contentBase64
        }, function (err, data) {
          if (err) {
            addMessage("assistant", "파일을 읽지 못했습니다: " + file.name + "\n" + err.message);
            setStatus("파일 읽기 실패");
            return;
          }
          attachments.push({ name: data.name || file.name, text: data.text || "" });
          addMessage("assistant", "소스 파일을 추가했습니다: " + (data.name || file.name));
          next();
        });
      };
      reader.onerror = function () {
        addMessage("assistant", "파일을 읽지 못했습니다: " + file.name);
        setStatus("파일 읽기 실패");
      };
      reader.readAsDataURL(file);
    }
    next();
  }

  document.getElementById("send").onclick = sendPrompt;
  document.getElementById("allowEdit").onclick = createPreview;
  commitPreviewEl.onclick = commitPreview;
  rollbackLastEl.onclick = rollbackLast;
  document.getElementById("selectAllChanges").onclick = function () {
    setAllChanges(true);
  };
  document.getElementById("clearAllChanges").onclick = function () {
    setAllChanges(false);
  };
  document.getElementById("cancelPlan").onclick = function () {
    hidePermission();
    addMessage("assistant", "편집을 취소했습니다. 요청을 다르게 쓰면 새 계획을 만들 수 있습니다.");
  };
  document.getElementById("refreshContext").onclick = refreshContext;
  document.getElementById("attachButton").onclick = function () {
    sourceFileEl.click();
  };
  if (clipboardImageButtonEl) {
    clipboardImageButtonEl.onclick = function () {
      importClipboardImageFromWindows(true);
    };
  }
  sourceFileEl.onchange = uploadFiles;
  attachmentListEl.onclick = function (e) {
    e = e || window.event;
    var target = e.target || e.srcElement;
    if (!target || target.className !== "attachmentRemove") return;
    var index = Number(target.getAttribute("data-index"));
    if (!isNaN(index) && index >= 0 && index < attachments.length) {
      attachments.splice(index, 1);
      updateAttachmentList();
    }
  };
  promptEl.onpaste = handlePaste;
  document.onpaste = handlePaste;
  promptEl.onfocus = refreshContext;
  promptEl.onkeydown = function (e) {
    e = e || window.event;
    if ((e.ctrlKey || e.metaKey) && e.keyCode === 86) {
      scheduleClipboardFallbackCheck();
      return true;
    }
    if (e.keyCode === 13 && !e.shiftKey && !e.isComposing) {
      sendPrompt();
      if (e.preventDefault) e.preventDefault();
      return false;
    }
    return true;
  };
  document.onkeydown = function (e) {
    e = e || window.event;
    var target = e.target || e.srcElement;
    if (target && target.id === "prompt") return true;
    if ((e.ctrlKey || e.metaKey) && e.keyCode === 86) {
      scheduleClipboardFallbackCheck();
    }
    return true;
  };
  window.onresize = resizeChat;

  refreshHealth();
  refreshContext();
  resizeChat();
})();
