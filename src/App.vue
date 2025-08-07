<!-- SPDX-License-Identifier: MIT -->
<script setup lang="ts">
import ControlBar from "@/components/ControlBar.vue";
import Ping from "ping.js";
import "bootstrap/dist/css/bootstrap.min.css";
import "bootstrap";
</script>

<script lang="ts">
export default {
  mounted() {
    const config_js = document.createElement("script");
    config_js.src = "./config.js";
    document.body.appendChild(config_js); // May fail. config.js is optional.
    const pre_soffice_js = document.createElement("script");
    pre_soffice_js.type = "module";
    pre_soffice_js.src = "./pre_soffice.js";
    pre_soffice_js.onload = function () {
      PingModule = Ping; // pass Ping module to plain JS
      ControlBar.init_control_bar();
    };
    document.body.appendChild(pre_soffice_js);
  },
};
</script>

<template>
  <div id="app">
    <div class="container-fluid p-0">
      <div class="row">
        <div class="col-12 mt-3">
          <h1>ZetaJS Ping Tool</h1>
        </div>
      </div>
      <div class="row">
        <div class="col-4 mt-3">
          <div class="input-group mb-3">
            <input
              type="text"
              id="ping_target"
              class="form-control"
              value="https://zetaoffice.net/"
              aria-label="Ping target"
              aria-describedby="btnPing"
            />
            <button class="btn btn-dark" type="button" id="btnPing" disabled>
              Ping
            </button>
          </div>
        </div>
      </div>
      <div class="row">
        <div class="col-4 mt-3">
          <div class="input-group mb-3">
            <label
              for="fileUpload"
              class="btn btn-outline-primary"
              id="lblUpload"
            >
              Choose File
            </label>
            <input
              type="file"
              id="fileUpload"
              class="form-control"
              accept=".ods,.xls,.xlsx"
              style="display: none"
            />
            <button
              class="btn btn-primary"
              type="button"
              id="btnUpload"
              disabled
            >
              Upload
            </button>
            <button
              class="btn btn-warning"
              type="button"
              id="btnDownloadODS"
              disabled
            >
              Download ODS
            </button>
            <button
              class="btn btn-info"
              type="button"
              id="btnDownloadXLS"
              disabled
            >
              Download XLS
            </button>
          </div>
        </div>
      </div>
      <div class="row">
        <div class="col-8">
          <ControlBar />
          <div
            onselectstart="event.preventDefault()"
            style="position: relative"
            class="mt-1"
          >
            <!--  position: Makes the loading animation overlay the canvas.
                    Needs a surrounding table with fixed width to work properly.
                  onselectstart: Prevents accidently selecting / highlighting the canvas.
                    Must be set on the surrounding HTML element. (tested in Firefox-128) -->
            <div
              id="loadingInfo"
              style="
                position: absolute;
                top: 50%;
                left: 50%;
                transform: translate(-50%, -50%);
              "
            >
              <div class="spinner"></div>
              <br />
              <h2>ZetaOffice is loading...</h2>
            </div>
            <canvas
              id="qtcanvas"
              contenteditable="true"
              oncontextmenu="event.preventDefault()"
              onkeydown="event.preventDefault()"
              style="
                border: 0px none;
                padding: 0;
                outline: 1px solid #cccccc;
                width: 1600px;
                height: 800px;
                visibility: hidden;
              "
            >
              <!-- QT requires the canvas to have the ID "qtcanvas". -->
              <!-- The canvas *must not* have any border or padding, or mouse coords will be wrong. -->
              <!-- An outline is fine though. -->
            </canvas>
          </div>
        </div>
      </div>
    </div>
  </div>
</template>

<style>
.spinner {
  border: 16px solid #1f2937; /* ZetaOffice brand color */
  border-top: 16px solid #059669; /* ZetaOffice brand color */
  border-radius: 50%;
  width: 120px;
  height: 120px;
  position: relative;
  left: 100px; /* adjust to center */
  animation: spin 2s linear 30; /* 60 seconds */
}
@keyframes spin {
  0% {
    transform: rotate(0deg);
  }
  100% {
    transform: rotate(360deg);
  }
}

/* Disable dark mode. ControlBar doesn't work well with it. */
body {
  background-color: var(--background-color);
  color: var(--text-color);
}
</style>
