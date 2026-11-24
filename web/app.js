/* PDF → Word: браузерная версия. Конвертация — в Web Worker (worker.js, Pyodide),
 * распознавание сканов — Tesseract.js. Файлы не покидают компьютер пользователя. */
(() => {
  "use strict";

  const $ = (sel, el = document) => el.querySelector(sel);
  const MAX_FILE_MB = 200;
  const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  const abs = (p) => new URL(p, location.href).href;

  const el = {
    hero: $("#hero"), workspace: $("#workspace"), grid: $("#fileGrid"), tpl: $("#cardTpl"),
    input: $("#fileInput"), pick: $("#pickBtn"), addMore: $("#addMoreBtn"), drop: $("#dropZone"),
    overlay: $("#dragOverlay"), convert: $("#convertBtn"), ocr: $("#ocrToggle"), ocrLang: $("#ocrLang"),
    langField: $("#langField"), title: $("#wsTitle"), banner: $("#doneBanner"), doneTitle: $("#doneTitle"),
    doneSub: $("#doneSub"), downloadAll: $("#downloadAllBtn"), restart: $("#restartBtn"), toasts: $("#toasts"),
    bootBox: $("#bootBox"), bootBar: $("#bootBar"), bootText: $("#bootText"),
    bootSide: $("#bootSide"), bootBarSide: $("#bootBarSide"), bootTextSide: $("#bootTextSide"), firstRun: $("#firstRun"),
  };

  const state = { files: [], ready: false, bootError: null, running: false };
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
  const toast = (text, isErr = false) => {
    const t = document.createElement("div");
    t.className = "toast" + (isErr ? " err" : "");
    t.textContent = text;
    el.toasts.appendChild(t);
    setTimeout(() => t.remove(), 5000);
  };
  const isPdf = (f) => /\.pdf$/i.test(f.name) || f.type === "application/pdf";
  const docxName = (name, ext = ".docx") => (name.replace(/\.pdf$/i, "").replace(/[\\/:*?"<>|]+/g, "_").trim() || "document") + ext;
  const saveBlob = (blob, name) => {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
  };

  // ------------------------------------------------------------------ рабочий поток с Python
  const worker = new Worker("worker.js");
  const pending = new Map();
  const listeners = new Map();   // key → функция прогресса

  worker.onmessage = (e) => {
    const m = e.data;
    if (m.event === "boot") return onBoot(m.pct, m.stage);
    if (m.event === "progress") return listeners.get(m.key)?.(m.pct, m.stage);
    const p = pending.get(m.id);
    if (!p) return;
    pending.delete(m.id);
    m.ok ? p.resolve(m.result) : p.reject(new Error(m.error));
  };
  worker.onerror = (e) => {
    onBootError("Не удалось запустить конвертер: " + (e.message || "ошибка загрузки"));
  };
  const rpc = (cmd, payload = {}, transfer = []) => new Promise((resolve, reject) => {
    const id = ++seq;
    pending.set(id, { resolve, reject });
    worker.postMessage({ id, cmd, ...payload }, transfer);
  });

  function onBoot(pct, stage) {
    for (const [bar, text] of [[el.bootBar, el.bootText], [el.bootBarSide, el.bootTextSide]]) {
      bar.style.width = `${pct}%`;
      text.textContent = pct >= 100 ? "Конвертер готов" : `${stage}… ${pct}%`;
    }
  }
  function onBootError(msg) {
    state.bootError = msg;
    for (const box of [el.bootBox, el.bootSide]) box.classList.add("is-error");
    el.bootText.textContent = el.bootTextSide.textContent = msg;
    el.bootSide.hidden = false;
    for (const it of state.files) {
      if (it.status === "reading" || it.status === "queued") { it.status = "error"; it.error = msg; paint(it); }
    }
    refresh();
  }

  const ready = (typeof WebAssembly === "object"
    ? rpc("init")
    : Promise.reject(new Error("Браузер не поддерживает WebAssembly — откройте страницу в свежем Chrome, Edge, Firefox или Яндекс Браузере.")))
    .then(() => {
      state.ready = true;
      for (const box of [el.bootBox, el.bootSide]) box.classList.add("is-ready");
      onBoot(100, "Готово");
      el.firstRun.hidden = true;
      refresh();
    })
    .catch((err) => onBootError(err.message));

  // ------------------------------------------------------------------ Tesseract.js (загружается только для сканов)
  let tess = null;
  function loadScript(src) {
    return new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = src; s.onload = resolve; s.onerror = () => reject(new Error("Не удалось загрузить модуль распознавания."));
      document.head.appendChild(s);
    });
  }
  async function ocrWorker(lang, onLog) {
    if (tess && tess.lang === lang) { tess.onLog = onLog; return tess.worker; }
    if (tess) { await tess.worker.terminate(); tess = null; }
    if (!window.Tesseract) await loadScript("tesseract/tesseract.min.js");
    const box = { lang, onLog, worker: null };
    box.worker = await window.Tesseract.createWorker(lang, 1, {
      workerPath: abs("tesseract/worker.min.js"),
      corePath: abs("tesseract/core/"),
      langPath: abs("tesseract/lang"),
      gzip: true,
      logger: (m) => box.onLog && box.onLog(m),
    });
    await box.worker.setParameters({ tessedit_pageseg_mode: "3", user_defined_dpi: "300" });
    tess = box;
    return box.worker;
  }

  // ------------------------------------------------------------------ файлы
  function addFiles(list) {
    let added = 0;
    for (const f of list) {
      if (!isPdf(f)) { toast(`«${f.name}» — не PDF. Можно добавлять только PDF-файлы.`, true); continue; }
      if (f.size > MAX_FILE_MB * 1024 * 1024) { toast(`«${f.name}» больше ${MAX_FILE_MB} МБ.`, true); continue; }
      if (f.size === 0) { toast(`«${f.name}» — пустой файл.`, true); continue; }
      const item = {
        key: `f${Date.now().toString(36)}${(++seq).toString(36)}`, file: f, name: f.name, size: f.size,
        status: "reading", info: null, thumbUrl: null, password: "", unlocked: false, pages: "",
        error: "", progress: 0, stage: "", blob: null, xlsx: null, url: null, stats: null,
      };
      state.files.push(item);
      mountCard(item);
      inspect(item);
      added++;
    }
    if (added) { showWorkspace(true); refresh(); }
  }

  async function inspect(item) {
    try {
      await ready;
      if (state.bootError) throw new Error(state.bootError);
      const buf = await item.file.arrayBuffer();
      const info = await rpc("inspect", { key: item.key, data: buf }, [buf]);
      item.file = null;  // содержимое теперь в памяти конвертера
      item.info = info;
      if (info.thumb) item.thumbUrl = URL.createObjectURL(new Blob([info.thumb], { type: "image/png" }));
      item.status = "ready";
    } catch (err) {
      item.status = "error";
      item.error = err.message;
      item.broken = true;
    }
    paint(item); refresh();
  }

  async function unlock(item) {
    const input = $(".pw input", item.el);
    const pw = input.value;
    if (!pw) { input.focus(); return; }
    $(".pw button", item.el).disabled = true;
    try {
      const res = await rpc("unlock", { key: item.key, password: pw });
      item.password = pw; item.unlocked = true; item.error = "";
      item.info.pages = res.pages;
      if (res.thumb) item.thumbUrl = URL.createObjectURL(new Blob([res.thumb], { type: "image/png" }));
    } catch (err) {
      item.error = err.message;
      input.select();
    }
    $(".pw button", item.el).disabled = false;
    paint(item); refresh();
  }

  function removeItem(item) {
    state.files = state.files.filter((i) => i !== item);
    item.el?.remove();
    if (item.thumbUrl) URL.revokeObjectURL(item.thumbUrl);
    if (item.url) URL.revokeObjectURL(item.url);
    listeners.delete(item.key);
    if (state.ready) rpc("drop", { key: item.key }).catch(() => {});
    if (!state.files.length) showWorkspace(false);
    refresh();
  }

  // ------------------------------------------------------------------ конвертация
  const convertible = (i) => (i.status === "ready" || (i.status === "error" && !i.broken))
    && !(i.info?.encrypted && !i.unlocked);

  function convertItems(items) {
    for (const item of items) {
      item.status = "queued"; item.error = ""; item.progress = 0; item.stage = "В очереди";
      paint(item);
    }
    refresh();
    runQueue();
  }

  async function runQueue() {
    if (state.running) return;
    state.running = true;
    try {
      let item;
      while ((item = state.files.find((i) => i.status === "queued"))) await convertOne(item);
    } finally {
      state.running = false;
      refresh();
    }
  }

  async function convertOne(item) {
    const setP = (pct, stage) => {
      item.progress = pct; item.stage = stage;
      paint(item);
    };
    item.status = "processing";
    setP(1, "Подготовка…");
    listeners.set(item.key, (pct, stage) => setP(Math.max(item.progress, pct), stage));
    const useOcr = el.ocr.checked;
    const lang = el.ocrLang.value || "rus+eng";
    try {
      const { ocr_pages: scans } = await rpc("start", {
        key: item.key,
        options: { password: item.password || null, pages: item.pages || null,
          ocr: useOcr ? "auto" : "off", ocr_lang: lang, title: item.name.replace(/\.pdf$/i, "") },
      });
      for (let k = 0; k < scans.length; k++) {
        const page = scans[k];
        const base = 5 + (40 * k) / scans.length, span = 40 / scans.length;
        setP(base, `Распознаю текст: страница ${page + 1}`);
        const png = await rpc("prepare", { key: item.key, page });
        try {
          const w = await ocrWorker(lang, (m) => {
            if (m.status === "recognizing text") setP(base + span * 0.2 + span * 0.8 * m.progress, `Распознаю текст: страница ${page + 1}`);
            else if (/loading|initializ/i.test(m.status)) setP(base, "Загружаю модуль распознавания…");
          });
          const r = await w.recognize(new Blob([png], { type: "image/png" }), {}, { hocr: true, text: false });
          await rpc("apply", { key: item.key, page, hocr: r.data.hocr });
        } catch (err) {
          console.warn("OCR failed", err);
          await rpc("skip", { key: item.key, page });
        }
      }
      const { docx, xlsx, stats } = await rpc("finish", { key: item.key });
      item.blob = new Blob([docx], { type: DOCX_MIME });
      item.xlsx = xlsx ? new Blob([xlsx], { type: XLSX_MIME }) : null;
      item.url = URL.createObjectURL(item.blob);
      item.stats = stats;
      item.status = "done";
    } catch (err) {
      item.status = "error";
      item.error = err.message;
    } finally {
      listeners.delete(item.key);
    }
    paint(item); refresh();
  }

  async function downloadAll() {
    const done = state.files.filter((i) => i.status === "done");
    if (done.length === 1) return saveBlob(done[0].blob, docxName(done[0].name));
    el.downloadAll.disabled = true;
    try {
      const files = await Promise.all(done.map(async (i) => ({ name: docxName(i.name), data: await i.blob.arrayBuffer() })));
      const zip = await rpc("zip", { files }, files.map((f) => f.data));
      saveBlob(new Blob([zip], { type: "application/zip" }), "pdf-to-word.zip");
    } catch (err) {
      toast("Не удалось собрать архив: " + err.message, true);
    } finally {
      el.downloadAll.disabled = false;
    }
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

  function paint(item) {
    const n = item.el;
    if (!n) return;
    const info = item.info || {};
    const st = item.status;
    n.className = "card" + (st === "done" ? " is-done" : "") + (st === "error" ? " is-error" : "")
      + (st === "queued" || st === "reading" ? " is-queued" : "");

    const img = $(".thumb img", n);
    if (item.thumbUrl && img.getAttribute("src") !== item.thumbUrl) {
      img.onload = () => { img.hidden = false; $(".thumb-ph", n).hidden = true; };
      img.src = item.thumbUrl;
    }

    const badges = $(".badges", n);
    badges.innerHTML = "";
    const badge = (text, cls) => {
      const b = document.createElement("span");
      b.className = "badge " + cls; b.textContent = text; badges.appendChild(b);
    };
    if (info.encrypted) badge(item.unlocked ? "🔓 Пароль" : "🔒 Пароль", "badge-lock");
    if (info.scanned_pages) badge(info.scanned_pages >= (info.pages || 1) ? "Скан" : `Сканов: ${info.scanned_pages}`, "badge-scan");

    const meta = $(".card-meta", n);
    if (st === "reading") {
      meta.textContent = state.ready ? "Открываю файл…" : "Ждёт готовности конвертера…";
    } else if (st === "done" && item.stats) {
      const s = item.stats;
      const parts = [`${s.pages_converted} стр. → DOCX ${fmtSize(item.blob.size)}`];
      if (s.tables) parts.push(`таблиц: ${s.tables}`);
      if (s.ocr_pages) parts.push(`OCR: ${s.ocr_pages} стр.`);
      if (s.headers) parts.push("шапка и штамп — в колонтитулах");
      meta.textContent = parts.join(" · ");
    } else {
      meta.textContent = [info.pages ? `${info.pages} стр.` : null, fmtSize(item.size)].filter(Boolean).join(" · ");
    }

    const needPw = info.encrypted && !item.unlocked && (st === "ready" || st === "error");
    $(".pw", n).hidden = !needPw;

    const editable = (st === "ready" || (st === "error" && !item.broken)) && !needPw;
    $(".pages-row", n).hidden = !(editable && (info.pages || 0) > 1);
    $(".pages-input", n).disabled = !editable;

    const ov = $(".overlay", n);
    const active = st === "queued" || st === "processing" || st === "reading";
    ov.hidden = !active;
    if (active) {
      const pct = Math.max(0, Math.min(100, item.progress || 0));
      const spinning = st !== "processing";
      $(".ring-fg", n).style.strokeDashoffset = spinning ? "" : String(119.4 * (1 - pct / 100));
      $(".ring-pct", n).textContent = spinning ? "…" : `${Math.round(pct)}%`;
      $(".overlay-stage", n).textContent = st === "reading" ? "Открываю файл" : (item.stage || "Обработка…");
    }
    $(".done-mark", n).hidden = st !== "done";

    const errBox = $(".card-error", n);
    errBox.hidden = !item.error;
    errBox.textContent = item.error || "";
    const infoBox = $(".card-info", n);
    const warns = (st === "done" && item.stats && item.stats.warnings) || [];
    infoBox.hidden = !warns.length;
    infoBox.textContent = warns.join(" ");

    const actions = $(".card-actions", n);
    actions.innerHTML = "";
    if (st === "done") {
      const b = document.createElement("button");
      b.className = "btn btn-ok"; b.type = "button"; b.textContent = "Скачать DOCX";
      b.addEventListener("click", () => saveBlob(item.blob, docxName(item.name)));
      actions.appendChild(b);
      if (item.xlsx) {
        const x = document.createElement("button");
        x.className = "btn btn-ghost btn-xlsx"; x.type = "button"; x.textContent = "Excel";
        x.title = "Таблицы и текст в Excel (.xlsx), лист на каждую страницу";
        x.addEventListener("click", () => saveBlob(item.xlsx, docxName(item.name, ".xlsx")));
        actions.appendChild(x);
      }
    } else if (st === "error" && !needPw && !item.broken) {
      const b = document.createElement("button");
      b.className = "btn btn-ghost"; b.type = "button"; b.textContent = "Повторить";
      b.addEventListener("click", () => convertItems([item]));
      actions.appendChild(b);
    }
  }

  function showWorkspace(on) {
    el.hero.hidden = on;
    el.workspace.hidden = !on;
    if (!on) el.banner.hidden = true;
    el.bootSide.hidden = !on || (state.ready && !state.bootError);
  }

  function refresh() {
    const files = state.files;
    const n = files.length;
    el.title.textContent = `${n} ${plural(n, ["файл", "файла", "файлов"])}`;
    el.bootSide.hidden = el.workspace.hidden || (state.ready && !state.bootError);

    const reading = files.some((i) => i.status === "reading");
    const active = files.some((i) => i.status === "queued" || i.status === "processing");
    const ready = files.filter(convertible);
    const done = files.filter((i) => i.status === "done");

    el.convert.disabled = !ready.length || active || !state.ready;
    el.convert.classList.toggle("is-finished", !ready.length && done.length > 0 && !active);
    const label = el.convert.firstChild;
    if (active) label.textContent = "Конвертация… ";
    else if (!state.ready && !state.bootError) label.textContent = "Подготовка конвертера… ";
    else if (reading && !ready.length) label.textContent = "Открываю файлы… ";
    else if (!ready.length && done.length) label.textContent = "✓ Всё сконвертировано ";
    else if (ready.length && ready.length < n && !done.length) label.textContent = `Конвертировать (${ready.length}) `;
    else label.textContent = "Конвертировать в Word ";

    const finished = n > 0 && !reading && !active && done.length > 0 && !files.some((i) => i.status === "ready");
    el.banner.hidden = !finished;
    if (finished) {
      const failed = files.filter((i) => i.status === "error").length;
      el.doneTitle.textContent = done.length === 1 ? "Готово! Документ Word создан" : `Готово! Документов Word: ${done.length}`;
      el.doneSub.textContent = failed
        ? `Не удалось сконвертировать: ${failed} ${plural(failed, ["файл", "файла", "файлов"])}.`
        : "Документы созданы в вашем браузере — скачайте их.";
      el.downloadAll.textContent = done.length === 1 ? "Скачать DOCX" : "Скачать все (ZIP)";
    }
  }

  // ------------------------------------------------------------------ события
  el.ocr.addEventListener("change", () => { el.langField.hidden = !el.ocr.checked; });
  el.pick.addEventListener("click", () => el.input.click());
  el.addMore.addEventListener("click", () => el.input.click());
  el.input.addEventListener("change", () => { addFiles([...el.input.files]); el.input.value = ""; });
  el.convert.addEventListener("click", () => convertItems(state.files.filter(convertible)));
  el.downloadAll.addEventListener("click", downloadAll);
  el.restart.addEventListener("click", () => {
    for (const item of [...state.files]) removeItem(item);
    showWorkspace(false);
    window.scrollTo({ top: 0, behavior: "smooth" });
  });

  let dragDepth = 0;
  const hasFiles = (e) => [...(e.dataTransfer?.types || [])].includes("Files");
  window.addEventListener("dragenter", (e) => {
    if (!hasFiles(e)) return;
    e.preventDefault(); dragDepth++;
    el.overlay.hidden = false; el.drop.classList.add("is-over");
  });
  window.addEventListener("dragover", (e) => { if (hasFiles(e)) e.preventDefault(); });
  window.addEventListener("dragleave", (e) => {
    if (!hasFiles(e)) return;
    dragDepth = Math.max(0, dragDepth - 1);
    if (!dragDepth) { el.overlay.hidden = true; el.drop.classList.remove("is-over"); }
  });
  window.addEventListener("drop", (e) => {
    if (!hasFiles(e)) return;
    e.preventDefault(); dragDepth = 0;
    el.overlay.hidden = true; el.drop.classList.remove("is-over");
    addFiles([...e.dataTransfer.files]);
  });
  window.addEventListener("beforeunload", (e) => {
    if (state.files.some((i) => ["queued", "processing"].includes(i.status))) { e.preventDefault(); e.returnValue = ""; }
  });
})();
