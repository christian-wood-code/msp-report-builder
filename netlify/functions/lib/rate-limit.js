"use strict";

// Small in-memory per-key limiter (resets on cold start; a brake, not a guarantee).
function makeLimiter(limit, windowMs) {
  const map = new Map();
  return function allow(key) {
    const now = Date.now();
    const e = map.get(key) || { count: 0, windowStart: now };
    if (now - e.windowStart > windowMs) { e.count = 0; e.windowStart = now; }
    e.count++;
    map.set(key, e);
    if (map.size > 500) for (const [k, v] of map) if (now - v.windowStart > windowMs) map.delete(k);
    return e.count <= limit;
  };
}

module.exports = { makeLimiter };
