//// IMPORTANT - Usage:
//// COPY THIS FILE to config.js to use it.




//////// SIMPLE OPTIONS

// Serve the LOWA files from a custom URL.
// E.g. for testing on localhost / 127.0.0.1.
//
// When serving the LOWA files from foreign origins, these HTTP headers are needed:
//   Cross-Origin-Resource-Policy: cross-origin
//   Access-Control-Allow-Origin: *
//   Access-Control-Allow-Methods: *
//
// IMPORTANT: Place soffice.{data,data.js.metadata,js,wasm} in this folder.
let config_soffice_base_url = '';
// current limitations:
// - Relative URLs may not work. Use absolute URLs instead!
//   - For example: http://127.0.0.1:8080/lowa_build_subdir/
// - Also always append the trailing slash!
