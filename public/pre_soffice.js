/* -*- Mode: JS; tab-width: 2; indent-tabs-mode: nil; js-indent-level: 2; fill-column: 100 -*- */
// SPDX-License-Identifier: MIT

import { ZetaHelperMain } from "./assets/vendor/zetajs/zetaHelper.js";

let thrPort; // zetajs thread communication
let tbDataJs; // toolbar dataset passed from vue.js for plain JS
window.PingModule = null; // Ping module passed from vue.js for plain JS

const loadingInfo = document.getElementById("loadingInfo");
const canvas = document.getElementById("qtcanvas");

// These elements will be retrieved when DOM is ready
let pingTarget,
  btnPing,
  btnExport,
  fileUpload,
  lblUpload,
  btnUpload,
  btnDownloadODS,
  btnDownloadXLS;
let disabledElementsAry = [];

// IMPORTANT:
// Set base URL to the soffice.* files.
// Use an empty string if those files are in the same directory.
let wasmPkg;
try {
  wasmPkg = "url:" + config_soffice_base_url; // May fail. config.js is optional.
} catch {}
const zHM = new ZetaHelperMain("office_thread.js", {
  threadJsType: "module",
  wasmPkg,
});

// Functions stored below window.* are usually accessed from vue.js.

window.jsPassCtrlBar = function (pTbDataJs) {
  tbDataJs = pTbDataJs;
  disabledElementsAry.push(tbDataJs);
};

window.toggleFormatting = function (id) {
  setToolbarActive(id, !tbDataJs.active[id]);
  thrPort.postMessage({ cmd: "toggleFormatting", id });
  // Give focus to the LO canvas to avoid issues with
  // <https://bugs.documentfoundation.org/show_bug.cgi?id=162291> "Setting Bold is
  // undone when clicking into non-empty document" when the user would need to click
  // into the canvas to give back focus to it:
  canvas.focus();
};

function setToolbarActive(id, value) {
  tbDataJs.active[id] = value;
  // Need to set "active" on "tbDataJs" to trigger an UI update.
  tbDataJs.active = tbDataJs.active;
}

let dbgPingData;
function pingResult(url, err, data) {
  dbgPingData = { data, err };
  const hostname = new URL(url).hostname;
  let output = data;
  // If /favicon.ico can't be loaded the result still represents the response time.
  if (err) output = hostname + ": " + output + " " + err;
  thrPort.postMessage({ cmd: "ping_result", id: { url, data } });
}

let pingInst;
const urls_ary = [
  "https://documentfoundation.org/",
  "https://ip4.me/",
  "https://allotropia.de/",
];
let urls_ary_i = 0;
function pingExamples(err, data) {
  let url = urls_ary[urls_ary_i];
  pingResult(url, err, data);
  url = urls_ary[++urls_ary_i];
  if (typeof url !== "undefined") {
    setTimeout(function () {
      // make the demo look more interesting ;-)
      pingInst.ping(url, function (err_rec, data_rec) {
        pingExamples(err_rec, data_rec);
      });
    }, 1000); // milliseconds
  }
}

function btnPingFunc() {
  // Using Ping callback interface.
  let url = pingTarget.value;
  if (!url.startsWith("http")) {
    url = "http://" + url;
  }
  pingInst.ping(url, function (err, data) {
    pingResult(url, err, data);
  });
}

// Define functions but don't set onclick handlers yet
function btnExportFunc() {
  thrPort.postMessage({ cmd: "export_excel" });
}

function btnUploadFunc() {
  if (!fileUpload || fileUpload.files.length === 0) return;

  for (const elem of disabledElementsAry) {
    if (elem) elem.disabled = true;
  }
  if (lblUpload) lblUpload.classList.add("disabled");

  const file = fileUpload.files[0];
  file.arrayBuffer().then((aryBuf) => {
    zHM.FS.writeFile("/tmp/uploaded_ping_data.ods", new Uint8Array(aryBuf));
    thrPort.postMessage({ cmd: "upload_file", filename: file.name });
  });
}

function btnDownloadODSFunc() {
  thrPort.postMessage({ cmd: "download_file", format: "ods" });
}

function btnDownloadXLSFunc() {
  thrPort.postMessage({ cmd: "download_file", format: "xls" });
}

async function get_calc_ping_example_ods() {
  const response = await fetch("./calc_ping_example.ods");
  return response.arrayBuffer();
}

zHM.start(function () {
  // Should run after App.vue has set PingModule but before demo().
  // 'Cross-Origin-Embedder-Policy': Ping seems to work with 'require-corp' without
  //   acutally having CORP on foreign origins.
  //   Also 'credentialless' isn't supported by Safari-18 as of 2024-09.
  pingInst = new window.PingModule();

  thrPort = zHM.thrPort;
  thrPort.onmessage = function (e) {
    switch (e.data.cmd) {
      case "setFormat":
        setToolbarActive(e.data.id, e.data.state);
        break;
      case "ui_ready":
        // Trigger resize of the embedded window to match the canvas size.
        // May somewhen be obsoleted by:
        //   https://gerrit.libreoffice.org/c/core/+/174040
        window.dispatchEvent(new Event("resize"));
        setTimeout(function () {
          // display Office UI properly
          loadingInfo.style.display = "none";
          canvas.style.visibility = null;

          // Get DOM elements now that Vue has rendered them
          pingTarget = document.getElementById("ping_target");
          btnPing = document.getElementById("btnPing");
          btnExport = document.getElementById("btnExport");
          fileUpload = document.getElementById("fileUpload");
          lblUpload = document.getElementById("lblUpload");
          btnUpload = document.getElementById("btnUpload");
          btnDownloadODS = document.getElementById("btnDownloadODS");
          btnDownloadXLS = document.getElementById("btnDownloadXLS");

          // Set up disabled elements array
          disabledElementsAry = [
            btnPing,
            btnExport,
            btnUpload,
            btnDownloadODS,
            btnDownloadXLS,
          ];

          // Enable all elements
          for (const elem of disabledElementsAry) {
            if (elem) elem.disabled = false;
          }

          // Keep upload disabled until file is selected
          if (btnUpload) btnUpload.disabled = true;

          // Set up all event handlers now that DOM elements exist
          if (btnPing) btnPing.onclick = btnPingFunc;
          if (btnExport) btnExport.onclick = btnExportFunc;
          if (btnUpload) btnUpload.onclick = btnUploadFunc;
          if (btnDownloadODS) btnDownloadODS.onclick = btnDownloadODSFunc;
          if (btnDownloadXLS) btnDownloadXLS.onclick = btnDownloadXLSFunc;

          // Upload functionality
          if (lblUpload) {
            lblUpload.onclick = function () {
              if (fileUpload) fileUpload.click();
            };
          }

          if (fileUpload) {
            fileUpload.onchange = function () {
              if (fileUpload.files.length > 0) {
                if (btnUpload) btnUpload.disabled = false;
                if (lblUpload) lblUpload.textContent = fileUpload.files[0].name;
              }
            };
          }

          if (pingTarget) {
            pingTarget.addEventListener("keyup", (evt) => {
              if (evt.key === "Enter") btnPingFunc();
            });
          }
          // Using Ping callback interface.
          pingInst.ping(urls_ary[urls_ary_i], function () {
            // Continue after first ping, which is often exceptionally slow.
            setTimeout(function () {
              // small delay to make the demo more interesting
              pingInst.ping(urls_ary[urls_ary_i], function (err, data) {
                pingExamples(err, data);
              });
            }, 1000); // milliseconds
          });
        }, 1000); // milliseconds
        break;
      case "export_complete":
        if (e.data.success) {
          // Read the exported file from virtual filesystem and create download
          // This follows the same pattern as the writer example
          try {
            const bytes = zHM.FS.readFile("/tmp/ping_export.xls");
            const blob = new Blob([bytes], {
              type: "application/vnd.ms-excel",
            });
            const link = document.createElement("a");
            link.href = URL.createObjectURL(blob);
            link.download = e.data.filename;
            link.style.display = "none";
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(link.href);
            alert(`✅ ${e.data.message} and download started!`);
          } catch (fileError) {
            alert(
              `✅ Export successful but download failed: ${fileError.toString()}`
            );
          }
        } else {
          alert(`❌ ${e.data.message}`);
        }
        break;
      case "upload_complete":
        if (e.data.success) {
          alert(`✅ File uploaded successfully: ${e.data.message}`);
          for (const elem of disabledElementsAry) {
            if (elem) elem.disabled = false;
          }
          if (lblUpload) {
            lblUpload.classList.remove("disabled");
            lblUpload.textContent = "Choose File";
          }
          if (fileUpload) fileUpload.value = "";
          if (btnUpload) btnUpload.disabled = true;
        } else {
          alert(`❌ Upload failed: ${e.data.message}`);
          for (const elem of disabledElementsAry) {
            if (elem) elem.disabled = false;
          }
          if (lblUpload) lblUpload.classList.remove("disabled");
          if (btnUpload) btnUpload.disabled = true;
        }
        break;
      case "download_complete":
        if (e.data.success) {
          try {
            const bytes = zHM.FS.readFile(e.data.tempPath);
            const mimeType =
              e.data.format === "ods"
                ? "application/vnd.oasis.opendocument.spreadsheet"
                : "application/vnd.ms-excel";
            const blob = new Blob([bytes], { type: mimeType });
            const link = document.createElement("a");
            link.href = URL.createObjectURL(blob);
            link.download = e.data.filename;
            link.style.display = "none";
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(link.href);
            alert(`✅ ${e.data.message} and download started!`);
          } catch (fileError) {
            alert(
              `✅ Export successful but download failed: ${fileError.toString()}`
            );
          }
        } else {
          alert(`❌ ${e.data.message}`);
        }
        break;
      default:
        throw Error("Unknown message command: " + e.data.cmd);
    }
  };

  get_calc_ping_example_ods().then(function (aryBuf) {
    zHM.FS.writeFile("/tmp/calc_ping_example.ods", new Uint8Array(aryBuf));
  });
});

/* vim:set shiftwidth=2 softtabstop=2 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */
