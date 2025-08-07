/* -*- Mode: JS; tab-width: 2; indent-tabs-mode: nil; js-indent-level: 2; fill-column: 100 -*- */
// SPDX-License-Identifier: MIT

// Debugging note:
// Switch the web worker in the browsers debug tab to debug this code.
// It's the "em-pthread" web worker with the most memory usage, where "zetajs" is defined.

// JS mode: module
import { ZetaHelperThread } from "./assets/vendor/zetajs/zetaHelper.js";

// global variables - zetajs environment:
const zHT = new ZetaHelperThread();
const zetajs = zHT.zetajs;
const css = zHT.css;
const desktop = zHT.desktop;

// = global variables =
// common variables:
let xModel, ctrl;
// export beans for file operations
let bean_overwrite, bean_excel_export, bean_ods_export;
// example specific:
let ping_line,
  xComponent,
  charLocale,
  formatNumber,
  formatText,
  activeSheet,
  cell;

// Export variables for debugging. Available for debugging via:
//   globalThis.zetajsStore.threadJsContext
export {
  zHT,
  xModel,
  ctrl,
  ping_line,
  xComponent,
  charLocale,
  formatNumber,
  formatText,
  activeSheet,
  cell,
  bean_overwrite,
  bean_excel_export,
  bean_ods_export,
};

function exportToExcel() {
  try {
    // Check if xModel is available
    if (!xModel) {
      throw new Error("Document model (xModel) is not available");
    }

    // Create a timestamp for the filename
    const timestamp = new Date()
      .toISOString()
      .replace(/[:.]/g, "-")
      .slice(0, -5);
    const filename = `ping_results_${timestamp}.xls`;

    // Use storeToURL to save the file to virtual filesystem
    // This follows the same pattern as the writer example
    xModel.storeToURL("file:///tmp/ping_export.xls", [
      bean_overwrite,
      bean_excel_export,
    ]);

    // Notify the main thread that export was successful
    zetajs.mainPort.postMessage({
      cmd: "export_complete",
      success: true,
      filename: filename,
      message: `Excel file exported successfully as ${filename}`,
    });
  } catch (error) {
    // Notify the main thread that export failed
    zetajs.mainPort.postMessage({
      cmd: "export_complete",
      success: false,
      error: error.toString(),
      message: `Export failed: ${error.toString()}`,
    });
  }
}

function uploadFile(filename) {
  try {
    // Close current document
    if (xModel) {
      xModel.close(true);
    }

    // Load the uploaded file
    xModel = desktop.loadComponentFromURL(
      "file:///tmp/uploaded_ping_data.ods",
      "_default",
      0,
      []
    );
    ctrl = xModel.getCurrentController();
    ctrl.getFrame().getContainerWindow().FullScreen = true;

    xComponent = ctrl.getModel();
    charLocale = xComponent.getPropertyValue("CharLocale");
    formatNumber = xComponent
      .getNumberFormats()
      .queryKey("0", charLocale, false);
    formatText = xComponent.getNumberFormats().queryKey("@", charLocale, false);

    activeSheet = ctrl.getActiveSheet();

    // Notify success
    zetajs.mainPort.postMessage({
      cmd: "upload_complete",
      success: true,
      message: `${filename} loaded successfully`,
    });
  } catch (error) {
    zetajs.mainPort.postMessage({
      cmd: "upload_complete",
      success: false,
      message: error.toString(),
    });
  }
}

function downloadFile(format) {
  try {
    if (!xModel) {
      throw new Error("No document available for download");
    }

    const timestamp = new Date()
      .toISOString()
      .replace(/[:.]/g, "-")
      .slice(0, -5);

    let filename, tempPath, bean_format;
    if (format === "ods") {
      filename = `ping_results_${timestamp}.ods`;
      tempPath = "/tmp/ping_download.ods";
      bean_format = bean_ods_export;
    } else {
      filename = `ping_results_${timestamp}.xls`;
      tempPath = "/tmp/ping_download.xls";
      bean_format = bean_excel_export;
    }

    // Export to virtual filesystem
    xModel.storeToURL(`file://${tempPath}`, [bean_overwrite, bean_format]);

    // Notify success
    zetajs.mainPort.postMessage({
      cmd: "download_complete",
      success: true,
      filename: filename,
      tempPath: tempPath,
      format: format,
      message: `${format.toUpperCase()} file exported successfully`,
    });
  } catch (error) {
    zetajs.mainPort.postMessage({
      cmd: "download_complete",
      success: false,
      error: error.toString(),
      message: `Download failed: ${error.toString()}`,
    });
  }
}

function demo() {
  // Initialize export beans similar to the writer example
  bean_overwrite = new css.beans.PropertyValue({
    Name: "Overwrite",
    Value: true,
  });
  bean_excel_export = new css.beans.PropertyValue({
    Name: "FilterName",
    Value: "MS Excel 97",
  });
  bean_ods_export = new css.beans.PropertyValue({
    Name: "FilterName",
    Value: "calc8",
  });

  zHT.configDisableToolbars(["Calc"]);

  xModel = desktop.loadComponentFromURL(
    "file:///tmp/calc_ping_example.ods",
    "_default",
    0,
    []
  );
  ctrl = xModel.getCurrentController();
  ctrl.getFrame().getContainerWindow().FullScreen = true;

  xComponent = ctrl.getModel();
  charLocale = xComponent.getPropertyValue("CharLocale");
  formatNumber = xComponent.getNumberFormats().queryKey("0", charLocale, false);
  formatText = xComponent.getNumberFormats().queryKey("@", charLocale, false);

  // Turn off UI elements:
  // zHT.dispatch(ctrl, 'Sidebar');
  // zHT.dispatch(ctrl, 'InputLineVisible');  // FormulaBar at the top
  // ctrl.getFrame().LayoutManager.hideElement("private:resource/statusbar/statusbar");
  // ctrl.getFrame().LayoutManager.hideElement("private:resource/menubar/menubar");

  for (const id of "Bold Italic Underline".split(" ")) {
    const urlObj = zHT.transformUrl(id);
    const listener = zetajs.unoObject([css.frame.XStatusListener], {
      disposing: function (source) {},
      statusChanged: function (state) {
        state = zetajs.fromAny(state.State);
        // Behave like desktop UI if a non uniformly formatted area is selected.
        if (typeof state !== "boolean") state = false; // like desktop UI
        zetajs.mainPort.postMessage({ cmd: "setFormat", id, state: state });
      },
    });
    zHT.queryDispatch(ctrl, urlObj).addStatusListener(listener, urlObj);
  }

  activeSheet = ctrl.getActiveSheet();

  // Add column headers
  let headerCell = activeSheet.getCellByPosition(0, 0);
  headerCell.setPropertyValue("NumberFormat", formatText);
  headerCell.setString("Hostname");

  headerCell = activeSheet.getCellByPosition(1, 0);
  headerCell.setPropertyValue("NumberFormat", formatText);
  headerCell.setString("Ping Time (ms)");

  zHT.thrPort.onmessage = function (e) {
    switch (e.data.cmd) {
      case "toggleFormatting":
        zHT.dispatch(ctrl, e.data.id);
        break;
      case "ping_result":
        if (ping_line === undefined) {
          ping_line = 1; // start at line 1 (line 0 is the header)
        } else {
          ping_line = findEmptyRowInCol1(activeSheet);
        }

        const url = e.data.id["url"];
        cell = activeSheet.getCellByPosition(0, ping_line);
        cell.setPropertyValue("NumberFormat", formatText); // optional
        cell.setString(new URL(url).hostname);

        cell = activeSheet.getCellByPosition(1, ping_line);
        let ping_value = String(e.data.id["data"]);
        if (!isNaN(ping_value)) {
          cell.setPropertyValue("NumberFormat", formatNumber); // optional
          cell.setValue(parseFloat(ping_value));
        } else {
          // in case e.data.id['data'] contains an error message
          cell.setPropertyValue("NumberFormat", formatText); // optional
          cell.setString(ping_value);
        }
        break;
      case "export_excel":
        exportToExcel();
        break;
      case "upload_file":
        uploadFile(e.data.filename);
        break;
      case "download_file":
        downloadFile(e.data.format);
        break;
      default:
        throw Error("Unknown message command: " + e.data.cmd);
    }
  };
  zHT.thrPort.postMessage({ cmd: "ui_ready" });
}

function findEmptyRowInCol1(activeSheet) {
  let str;
  let line = 0;
  while (str != "") {
    line++;
    str = activeSheet.getCellByPosition(0, line).getString();
  }
  return line;
}

demo(); // launching demo

/* vim:set shiftwidth=2 softtabstop=2 expandtab cinoptions=b1,g0,N-s cinkeys+=0=break: */
