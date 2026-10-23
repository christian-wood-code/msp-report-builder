"use strict";

// Size / shape limits for the export endpoints. Real report payloads are ~130 KB (logos included);
// the cap leaves headroom but stops anyone feeding the renderer something huge or pathological.
const MAX_BYTES = 600 * 1024;
const MAX_ARRAY = 5000;
const MAX_DEPTH = 12;

function tooComplex(v, depth) {
  if (depth > MAX_DEPTH) return true;
  if (Array.isArray(v)) {
    if (v.length > MAX_ARRAY) return true;
    for (const x of v) if (x && typeof x === 'object' && tooComplex(x, depth + 1)) return true;
  } else if (v && typeof v === 'object') {
    for (const k of Object.keys(v)) {
      const x = v[k];
      if (x && typeof x === 'object' && tooComplex(x, depth + 1)) return true;
    }
  }
  return false;
}

// Each returns an error message, or null when acceptable.
function checkBodySize(rawBody) {
  return rawBody && rawBody.length > MAX_BYTES ? 'Payload too large' : null;
}
function checkShape(parsed) {
  return tooComplex(parsed, 0) ? 'Payload too complex' : null;
}

module.exports = { checkBodySize, checkShape, MAX_BYTES };
