/* Web Worker: Python-конвейер конвертации (тот же код, что на сервере) внутри Pyodide.
 *
 * Сообщения от страницы:  {id, cmd, ...}  →  ответ {id, ok, result | error}
 *   init                         загрузить Pyodide, пакеты и код конвертера
 *   inspect  {key, data}         страницы, пароль, сканы, превью
 *   unlock   {key, password}     проверить пароль, получить превью
 *   start    {key, options}      начать конвертацию → список страниц-сканов для OCR
 *   prepare  {key, page}         подготовить скан → PNG для Tesseract.js
 *   apply    {key, page, hocr}   собрать страницу из hOCR
 *   skip     {key, page}         распознать не удалось — оставить страницу картинкой
 *   finish   {key}               собрать DOCX
 *   zip      {files:[{name,data}]}
 *   drop     {key}               освободить память
 * Прогресс: {event:"progress", key, pct, stage}; загрузка: {event:"boot", pct, stage}
 */
"use strict";

let py = null;
let manifest = null;

const post = (msg, transfer) => self.postMessage(msg, transfer || []);
const boot = (pct, stage) => post({ event: "boot", pct, stage });

async function fetchBytes(url) {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Не удалось загрузить ${url} (${res.status})`);
  return new Uint8Array(await res.arrayBuffer());
}

async function init() {
  if (py) return;
  boot(2, "Загружаю описание сборки");
  manifest = await (await fetch("manifest.json")).json();

  boot(5, "Загружаю Python (WebAssembly)");
  importScripts("pyodide/pyodide.js");
  py = await loadPyodide({ indexURL: new URL("pyodide/", self.location.href).href });

  const pkgs = manifest.pyodide_packages;
  let done = 0;
  boot(25, "Загружаю библиотеки PDF и Word");
  await py.loadPackage(pkgs, {
    messageCallback: (m) => {
      if (/^Loaded /.test(m)) {
        done += 1;
        boot(25 + Math.round((60 * done) / pkgs.length), "Загружаю библиотеки PDF и Word");
      }
    },
    errorCallback: () => {},
  });

  boot(86, "Устанавливаю конвертер");
  for (const w of manifest.wheels) py.unpackArchive(await fetchBytes(w), "wheel");
  py.unpackArchive(await fetchBytes(manifest.app_zip), "zip", { extractDir: "/home/pyodide" });
  py.FS.mkdirTree("/fonts");
  for (const [name, url] of Object.entries(manifest.fonts)) py.FS.writeFile(`/fonts/${name}`, await fetchBytes(url));
  py.FS.mkdirTree("/work");

  boot(94, "Запускаю конвертер");
  self.pyProgress = (key, pct, stage) => post({ event: "progress", key, pct, stage });
  await py.runPythonAsync(`
import os, sys, json, logging, warnings
os.environ["OCR_FONT"] = "/fonts/LiberationSerif-Regular.ttf"
os.environ["OCR_FONT_BOLD"] = "/fonts/LiberationSerif-Bold.ttf"
os.environ["DATA_DIR"] = "/work"
warnings.filterwarnings("ignore")
logging.basicConfig(level=logging.ERROR)
sys.path.insert(0, "/home/pyodide")
from app import webapi
webapi.warmup()
`);
  boot(100, "Готово");
}

function call(fn, args) {
  // аргументы передаём через глобальную переменную — без экранирования строк
  py.globals.set("_ARGS", py.toPy(args));
  try {
    const res = py.runPython(`webapi.${fn}(**_ARGS)`);
    return res && res.toJs ? res.toJs({ dict_converter: Object.fromEntries, create_proxies: false }) : res;
  } finally {
    py.globals.delete("_ARGS");
  }
}

function pyError(err) {
  const text = String(err && err.message ? err.message : err);
  const m = text.match(/(?:ConversionError|OcrError): (.*)/);
  if (m) return m[1].trim();
  console.error(text);
  return "Не удалось обработать документ: внутренняя ошибка конвертации.";
}

self.onmessage = async (e) => {
  const { id, cmd } = e.data;
  try {
    let result;
    const transfer = [];
    switch (cmd) {
      case "init":
        await init();
        result = { version: manifest.version };
        break;
      case "inspect":
        py.FS.writeFile(`/work/${e.data.key}.pdf`, new Uint8Array(e.data.data));
        result = call("inspect", { key: e.data.key });
        break;
      case "unlock":
        result = call("unlock", { key: e.data.key, password: e.data.password });
        break;
      case "start":
        result = call("start", { key: e.data.key, options: e.data.options });
        break;
      case "prepare":
        result = call("prepare", { key: e.data.key, page: e.data.page });
        transfer.push(result.buffer);
        break;
      case "apply":
        result = call("apply", { key: e.data.key, page: e.data.page, hocr: e.data.hocr });
        break;
      case "skip":
        result = call("skip", { key: e.data.key, page: e.data.page });
        break;
      case "finish":
        result = call("finish", { key: e.data.key });
        transfer.push(result.docx.buffer);
        break;
      case "zip":
        result = call("make_zip", { files: e.data.files.map((f) => [f.name, new Uint8Array(f.data)]) });
        transfer.push(result.buffer);
        break;
      case "drop":
        result = call("drop", { key: e.data.key });
        break;
      default:
        throw new Error(`unknown command ${cmd}`);
    }
    post({ id, ok: true, result }, transfer);
  } catch (err) {
    post({ id, ok: false, error: pyError(err) });
  }
};
