(() => {
  "use strict";

  const $ = (sel, el = document) => el.querySelector(sel);
  const MAX_PARALLEL_UPLOADS = 3;
  const POLL_MS = 800;

  const el = {
    hero: $("#hero"), workspace: $("#workspace"), grid: $("#fileGrid"), tpl: $("#cardTpl"),
    input: $("#fileInput"), pick: $("#pickBtn"), addMore: $("#addMoreBtn"), drop: $("#dropZone"),
    overlay: $("#dragOverlay"), convert: $("#convertBtn"), ocr: $("#ocrToggle"), ocrLang: $("#ocrLang"),
    langField: $("#langField"), ocrHelp: $("#ocrHelp"), title: $("#wsTitle"), limits: $("#limits"),
    banner: $("#doneBanner"), doneTitle: $("#doneTitle"), doneSub: $("#doneSub"),
    downloadAll: $("#downloadAllBtn"), restart: $("#restartBtn"), toasts: $("#toasts"),
    ttlPill: $("#ttlPill"), ttlFeature: $("#ttlFeature"), sideFoot: $("#sideFoot"),
  };

  const state = { cfg: null, files: [], uploading: 0, timer: null };
  let seq = 0;

  // ------------------------------------------------------------------ утилиты
  const plural = (n, forms) => {
    const a = Math.abs(n) % 100, b = a % 10;
    if (a > 10 && a < 20) return forms[2];
    if (b > 1 && b < 5) return forms[1];
    if (b === 1) return forms[0];
    return forms[2];
  };
  const fmtSize = (bytes) => {
    if (bytes == null) return "";
    if (bytes < 1024) return `${bytes} Б`;
    if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} КБ`;
    return `${(bytes / 1024 / 1024).toFixed(bytes < 10 * 1024 * 1024 ? 1 : 0).replace(".", ",")} МБ`;
  };
  const errMsg = (body, fallback = "Что-то пошло не так. Попробуйте ещё раз.") => {
    const d = body && body.detail;
    if (typeof d === "string") return d;
    if (Array.isArray(d)) return "Проверьте параметры: диапазон страниц или язык указаны неверно.";
    return fallback;
  };
  const toast = (text, isErr = false) => {
    const t = document.createElement("div");
    t.className = "toast" + (isErr ? " err" : "");
    t.textContent = text;
    el.toasts.appendChild(t);
    setTimeout(() => t.remove(), 5000);
  };
  const api = async (method, url, body) => {
    const res = await fetch(url, {
      method,
      headers: body ? { "Content-Type": "application/json" } : undefined,
      body: body ? JSON.stringify(body) : undefined,
    });
    let data = null;
    try { data = await res.json(); } catch { /* пустой ответ */ }
    return { ok: res.ok, status: res.status, data };
  };
  const isPdf = (f) => /\.pdf$/i.test(f.name) || f.type === "application/pdf";

  // ------------------------------------------------------------------ конфигурация
  async function loadConfig() {
    try {
      const { ok, data } = await api("GET", "api/config");
      if (!ok) throw new Error();
      state.cfg = data;
    } catch {
      state.cfg = { max_file_mb: 100, file_ttl_min: 60, ocr: { available: false, languages: [] } };
    }
    const c = state.cfg;
    if (c.notice) {
      const n = $("#notice");
      n.textContent = c.notice;
      n.hidden = false;
    }
    el.limits.textContent = `PDF до ${c.max_file_mb} МБ` + (c.max_pages ? ` и ${c.max_pages} страниц` : "");
    el.ttlPill.textContent = `Файлы хранятся ${c.file_ttl_min} мин`;
    el.ttlPill.hidden = false;
    el.ttlFeature.textContent = `Файлы удаляются с сервера через ${c.file_ttl_min} мин`;
    el.sideFoot.textContent = `Файлы будут удалены с сервера через ${c.file_ttl_min} минут.`;

    el.ocrLang.innerHTML = "";
    const langs = c.ocr.languages || [];
    const opts = [];
    const has = (code) => langs.some((l) => l.code === code);
    if (has("rus") && has("eng")) opts.push(["rus+eng", "Русский + English"]);
    langs.forEach((l) => opts.push([l.code, l.name]));
    for (const [code, name] of opts) {
      const o = document.createElement("option");
      o.value = code; o.textContent = name;
      if (code === c.ocr.default) o.selected = true;
      el.ocrLang.appendChild(o);
    }
    if (!c.ocr.available) {
      el.ocr.checked = false;
      el.ocr.disabled = true;
      el.langField.hidden = true;
      el.ocrHelp.textContent = "Распознавание сканов на этом сервере отключено: отсканированные страницы будут вставлены в Word картинками.";
    }
  }
  el.ocr.addEventListener("change", () => { el.langField.hidden = !el.ocr.checked; });

  // ------------------------------------------------------------------ добавление файлов
  function addFiles(list) {
    const maxBytes = (state.cfg?.max_file_mb || 100) * 1024 * 1024;
    let added = 0;
    for (const f of list) {
      if (!isPdf(f)) { toast(`«${f.name}» — не PDF. Можно добавлять только PDF-файлы.`, true); continue; }
      if (f.size > maxBytes) { toast(`«${f.name}» больше ${state.cfg.max_file_mb} МБ.`, true); continue; }
      if (f.size === 0) { toast(`«${f.name}» — пустой файл.`, true); continue; }
      const item = {
        lid: ++seq, file: f, name: f.name, size: f.size, status: "wait-upload",
        uploadPct: 0, task: null, password: "", unlocked: false, pages: "", error: "", xhr: null,
      };
      state.files.push(item);
      mountCard(item);
      added++;
    }
    if (added) {
      showWorkspace(true);
      pumpUploads();
      refresh();
    }
  }

  function pumpUploads() {
    while (state.uploading < MAX_PARALLEL_UPLOADS) {
      const next = state.files.find((i) => i.status === "wait-upload");
      if (!next) break;
      upload(next);
    }
  }

  function upload(item) {
    item.status = "uploading";
    state.uploading++;
    const xhr = new XMLHttpRequest();
    item.xhr = xhr;
    xhr.open("POST", "api/tasks");
    xhr.responseType = "json";
    xhr.upload.onprogress = (e) => {
      if (e.lengthComputable) { item.uploadPct = e.loaded / e.total; paint(item); }
    };
    const finish = () => { state.uploading--; item.xhr = null; paint(item); pumpUploads(); refresh(); };
    xhr.onload = () => {
      if (xhr.status === 201 && xhr.response) {
        item.task = xhr.response;
        item.status = "ready";
        item.error = "";
      } else {
        item.status = "upload-error";
        item.error = errMsg(xhr.response, `Ошибка загрузки (${xhr.status}).`);
      }
      finish();
    };
    xhr.onerror = () => { item.status = "upload-error"; item.error = "Не удалось загрузить: нет связи с сервером."; finish(); };
    xhr.onabort = () => { state.uploading--; pumpUploads(); refresh(); };
    const form = new FormData();
    form.append("file", item.file, item.file.name);
    xhr.send(form);
    paint(item);
  }

  // ------------------------------------------------------------------ карточки
  function mountCard(item) {
    const node = el.tpl.content.firstElementChild.cloneNode(true);
    item.el = node;
    $(".card-name", node).textContent = item.name;
    $(".card-name", node).title = item.name;
    $(".card-remove", node).addEventListener("click", () => removeItem(item));
    $(".pw", node).addEventListener("submit", (e) => { e.preventDefault(); unlock(item); });
    const pagesInput = $(".pages-input", node);
    pagesInput.addEventListener("input", () => {
      pagesInput.value = pagesInput.value.replace(/[^\d,\-–;\s]/g, "");
      item.pages = pagesInput.value.trim();
    });
    el.grid.appendChild(node);
    paint(item);
  }

  function setThumb(item) {
    const img = $(".thumb img", item.el);
    if (item.task && item.task.has_thumb) {
      const src = `api/tasks/${item.task.id}/thumb?v=${item.unlocked ? 1 : 0}`;
      if (img.getAttribute("src") !== src) {
        img.onload = () => { img.hidden = false; $(".thumb-ph", item.el).hidden = true; };
        img.src = src;
      }
    }
  }

  function paint(item) {
    const n = item.el;
    if (!n) return;
    const t = item.task || {};
    const st = item.status;
    n.className = "card" + (st === "done" ? " is-done" : "") + (st === "error" || st === "upload-error" ? " is-error" : "")
      + (st === "queued" ? " is-queued" : "");

    setThumb(item);

    // значки
    const badges = $(".badges", n);
    badges.innerHTML = "";
    const badge = (text, cls) => {
      const b = document.createElement("span");
      b.className = "badge " + cls; b.textContent = text; badges.appendChild(b);
    };
    if (t.encrypted) badge(item.unlocked ? "🔓 Пароль" : "🔒 Пароль", "badge-lock");
    if (t.scanned_pages) badge(t.scanned_pages >= (t.pages || 1) ? "Скан" : `Сканов: ${t.scanned_pages}`, "badge-scan");

    // описание
    const meta = $(".card-meta", n);
    if (st === "uploading" || st === "wait-upload") {
      meta.textContent = st === "wait-upload" ? "Ожидает загрузки…" : `Загрузка… ${Math.round(item.uploadPct * 100)}%`;
    } else if (st === "done" && t.stats) {
      const s = t.stats;
      const parts = [`${s.pages_converted} ${plural(s.pages_converted, ["стр.", "стр.", "стр."])} → DOCX ${fmtSize(t.output_size)}`];
      if (s.tables) parts.push(`таблиц: ${s.tables}`);
      if (s.ocr_pages) parts.push(`OCR: ${s.ocr_pages} стр.`);
      if (s.headers) parts.push("шапка и штамп — в колонтитулах");
      meta.textContent = parts.join(" · ");
    } else {
      meta.textContent = [t.pages ? `${t.pages} стр.` : null, fmtSize(item.size)].filter(Boolean).join(" · ");
    }

    const bar = $(".upload-bar", n);
    bar.hidden = st !== "uploading";
    $("i", bar).style.width = `${Math.round(item.uploadPct * 100)}%`;

    // пароль
    const needPw = t.encrypted && !item.unlocked && (st === "ready" || st === "error");
    $(".pw", n).hidden = !needPw;

    // диапазон страниц
    const editable = (st === "ready" || st === "error") && !needPw;
    $(".pages-row", n).hidden = !(editable && (t.pages || 0) > 1);
    $(".pages-input", n).disabled = !editable;

    // прогресс
    const ov = $(".overlay", n);
    const active = st === "queued" || st === "processing";
    ov.hidden = !active;
    if (active) {
      const pct = Math.max(0, Math.min(100, t.progress || 0));
      $(".ring-fg", n).style.strokeDashoffset = st === "queued" ? "" : String(119.4 * (1 - pct / 100));
      $(".ring-pct", n).textContent = st === "queued" ? "…" : `${Math.round(pct)}%`;
      $(".overlay-stage", n).textContent = t.stage || (st === "queued" ? "В очереди" : "Обработка…");
    }
    $(".done-mark", n).hidden = st !== "done";

    // ошибки и предупреждения
    const errBox = $(".card-error", n);
    errBox.hidden = !item.error;
    errBox.textContent = item.error || "";
    const info = $(".card-info", n);
    const warns = (st === "done" && t.stats && t.stats.warnings) || [];
    info.hidden = !warns.length;
    info.textContent = warns.join(" ");

    // кнопки
    const actions = $(".card-actions", n);
    actions.innerHTML = "";
    if (st === "done") {
      const a = document.createElement("a");
      a.className = "btn btn-ok";
      a.href = `api/tasks/${t.id}/download`;
      a.textContent = "Скачать DOCX";
      actions.appendChild(a);
      if (t.has_xlsx) {
        const x = document.createElement("a");
        x.className = "btn btn-ghost btn-xlsx";
        x.href = `api/tasks/${t.id}/download?format=xlsx`;
        x.textContent = "Excel";
        x.title = "Таблицы и текст в Excel (.xlsx), лист на каждую страницу";
        actions.appendChild(x);
      }
    } else if (st === "error" && !needPw) {
      const b = document.createElement("button");
      b.className = "btn btn-ghost"; b.type = "button"; b.textContent = "Повторить";
      b.addEventListener("click", () => convertItems([item]));
      actions.appendChild(b);
    } else if (st === "upload-error") {
      const b = document.createElement("button");
      b.className = "btn btn-ghost"; b.type = "button"; b.textContent = "Загрузить снова";
      b.addEventListener("click", () => { item.status = "wait-upload"; item.error = ""; item.uploadPct = 0; pumpUploads(); refresh(); });
      actions.appendChild(b);
    }
  }

  async function unlock(item) {
    const input = $(".pw input", item.el);
    const pw = input.value;
    if (!pw) { input.focus(); return; }
    const btn = $(".pw button", item.el);
    btn.disabled = true;
    const { ok, data } = await api("POST", `api/tasks/${item.task.id}/unlock`, { password: pw });
    btn.disabled = false;
    if (ok) {
      item.password = pw;
      item.unlocked = true;
      item.task = data;
      item.error = "";
    } else {
      item.error = errMsg(data);
      input.select();
    }
    paint(item);
    refresh();
  }

  async function removeItem(item) {
    if (item.xhr) item.xhr.abort();
    state.files = state.files.filter((i) => i !== item);
    item.el?.remove();
    if (item.task) api("DELETE", `api/tasks/${item.task.id}`).catch(() => {});
    if (!state.files.length) showWorkspace(false);
    refresh();
  }

  // ------------------------------------------------------------------ конвертация
  const convertible = (i) => (i.status === "ready" || i.status === "error") && i.task && !(i.task.encrypted && !i.unlocked);

  async function convertItems(items) {
    const options = {
      ocr: el.ocr.checked ? "auto" : "off",
      ocr_lang: el.ocr.checked ? el.ocrLang.value || null : null,
    };
    await Promise.all(items.map(async (item) => {
      item.status = "queued";
      item.error = "";
      item.task.progress = 0;
      item.task.stage = "В очереди";
      paint(item);
      const body = { ...options, pages: item.pages || null, password: item.password || null };
      const { ok, data } = await api("POST", `api/tasks/${item.task.id}/convert`, body);
      if (ok) {
        item.task = data;
        item.status = data.status;
      } else {
        item.status = "error";
        item.error = errMsg(data);
      }
      paint(item);
    }));
    refresh();
    startPolling();
  }

  function startPolling() {
    if (state.timer) return;
    state.timer = setInterval(poll, POLL_MS);
  }

  async function poll() {
    const active = state.files.filter((i) => i.status === "queued" || i.status === "processing");
    if (!active.length) { clearInterval(state.timer); state.timer = null; return; }
    await Promise.all(active.map(async (item) => {
      try {
        const { ok, status, data } = await api("GET", `api/tasks/${item.task.id}`);
        if (ok) {
          item.task = data;
          item.status = data.status;
          item.error = data.status === "error" ? (data.error || "Ошибка конвертации.") : "";
        } else if (status === 404) {
          item.status = "upload-error";
          item.error = "Файл удалён с сервера. Загрузите его снова.";
        }
      } catch { /* временная ошибка сети — повторим */ }
      paint(item);
    }));
    refresh();
  }

  // ------------------------------------------------------------------ общее состояние
  function showWorkspace(on) {
    el.hero.hidden = on;
    el.workspace.hidden = !on;
    if (!on) el.banner.hidden = true;
  }

  function refresh() {
    const files = state.files;
    const n = files.length;
    el.title.textContent = `${n} ${plural(n, ["файл", "файла", "файлов"])}`;

    const uploading = files.some((i) => i.status === "uploading" || i.status === "wait-upload");
    const active = files.some((i) => i.status === "queued" || i.status === "processing");
    const ready = files.filter(convertible);
    const locked = files.filter((i) => i.task && i.task.encrypted && !i.unlocked && i.status === "ready");
    const done = files.filter((i) => i.status === "done");

    el.convert.disabled = !ready.length || active;
    el.convert.classList.toggle("is-finished", !ready.length && done.length > 0 && !active);
    const label = el.convert.firstChild;
    if (active) label.textContent = "Конвертация… ";
    else if (uploading && !ready.length) label.textContent = "Загрузка файлов… ";
    else if (!ready.length && done.length) label.textContent = "✓ Всё сконвертировано ";
    else if (ready.length && ready.length < n && !done.length) label.textContent = `Конвертировать (${ready.length}) `;
    else label.textContent = "Конвертировать в Word ";

    if (locked.length && !active) {
      el.sideFoot.textContent = `Введите пароль для ${locked.length} ${plural(locked.length, ["файла", "файлов", "файлов"])} на карточке.`;
    } else if (state.cfg) {
      el.sideFoot.textContent = `Файлы будут удалены с сервера через ${state.cfg.file_ttl_min} минут.`;
    }

    const finished = n > 0 && !uploading && !active && done.length > 0 && !ready.some((i) => i.status === "ready");
    el.banner.hidden = !finished;
    if (finished) {
      const failed = files.filter((i) => i.status === "error" || i.status === "upload-error").length;
      el.doneTitle.textContent = done.length === 1 ? "Готово! Документ Word создан" : `Готово! Документов Word: ${done.length}`;
      el.doneSub.textContent = failed
        ? `Не удалось сконвертировать: ${failed} ${plural(failed, ["файл", "файла", "файлов"])}.`
        : "Скачайте результат — файлы будут автоматически удалены с сервера.";
      if (done.length === 1) {
        el.downloadAll.href = `api/tasks/${done[0].task.id}/download`;
        el.downloadAll.textContent = "Скачать DOCX";
      } else {
        el.downloadAll.href = `api/zip?ids=${done.map((i) => i.task.id).join(",")}`;
        el.downloadAll.textContent = "Скачать все (ZIP)";
      }
    }
  }

  // ------------------------------------------------------------------ события
  el.pick.addEventListener("click", () => el.input.click());
  el.addMore.addEventListener("click", () => el.input.click());
  el.input.addEventListener("change", () => { addFiles([...el.input.files]); el.input.value = ""; });
  el.convert.addEventListener("click", () => convertItems(state.files.filter(convertible)));
  el.restart.addEventListener("click", () => {
    for (const item of [...state.files]) removeItem(item);
    showWorkspace(false);
    window.scrollTo({ top: 0, behavior: "smooth" });
  });

  let dragDepth = 0;
  const hasFiles = (e) => [...(e.dataTransfer?.types || [])].includes("Files");
  window.addEventListener("dragenter", (e) => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    dragDepth++;
    el.overlay.hidden = false;
    el.drop.classList.add("is-over");
  });
  window.addEventListener("dragover", (e) => { if (hasFiles(e)) e.preventDefault(); });
  window.addEventListener("dragleave", (e) => {
    if (!hasFiles(e)) return;
    dragDepth = Math.max(0, dragDepth - 1);
    if (!dragDepth) { el.overlay.hidden = true; el.drop.classList.remove("is-over"); }
  });
  window.addEventListener("drop", (e) => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    dragDepth = 0;
    el.overlay.hidden = true;
    el.drop.classList.remove("is-over");
    addFiles([...e.dataTransfer.files]);
  });

  window.addEventListener("beforeunload", (e) => {
    const busy = state.files.some((i) => ["uploading", "wait-upload", "queued", "processing"].includes(i.status));
    if (busy) { e.preventDefault(); e.returnValue = ""; }
  });

  loadConfig();
})();
