"use strict";
const crypto = require("crypto");

// Constant-time string comparison for shared secrets. Both values are hashed first so the two
// buffers are always the same length (timingSafeEqual throws on different lengths) and the
// comparison time does not depend on how many leading characters match.
function safeEqual(a, b) {
  const ha = crypto.createHash("sha256").update(String(a == null ? "" : a)).digest();
  const hb = crypto.createHash("sha256").update(String(b == null ? "" : b)).digest();
  return crypto.timingSafeEqual(ha, hb);
}

module.exports = { safeEqual };
