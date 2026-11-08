/* Muir Receipts — receipt "scanner": find the receipt in a photo, straighten it, and boost the text.
   No libraries: everything runs on the phone with <canvas>, so it works offline and costs no download.
   window.MuirScan = { analyse, render, ... }
     analyse(img)  → { quad:[[x,y]×4 TL,TR,BR,BL in image px], rot, method, confident }
     render(img, { quad, rot, enhance, maxEdge, quality }) → Promise<Blob> (JPEG)
   Detection, in order:
     1. paper vs background (receipt on a darker or coloured surface) — brightness/whiteness split
     2. the receipt's two long edges (white receipt on a white desk) — straight-line search
     3. nothing confident → the whole photo (the person drags the corners)
   Enhance = remove shadows/uneven light (divide by a blurred background), then stretch contrast. */
(() => {
  "use strict";
  const WORK = 420;                       // long edge of the small copy used for detection

  // ---------- small helpers ----------
  function smallCopy(img) {
    const iw = img.naturalWidth || img.width, ih = img.naturalHeight || img.height;
    const s = Math.min(1, WORK / Math.max(iw, ih));
    const w = Math.max(8, Math.round(iw * s)), h = Math.max(8, Math.round(ih * s));
    const c = document.createElement("canvas"); c.width = w; c.height = h;
    const x = c.getContext("2d", { willReadFrequently: true }); x.drawImage(img, 0, 0, w, h);
    return { w, h, s, data: x.getImageData(0, 0, w, h).data };
  }
  function otsu(vals, lo = 0, hi = 255) {
    const bins = 256, hist = new Float64Array(bins), span = (hi - lo) || 1;
    for (let i = 0; i < vals.length; i++) hist[Math.max(0, Math.min(255, Math.round((vals[i] - lo) / span * 255)))]++;
    let sum = 0; for (let i = 0; i < bins; i++) sum += i * hist[i];
    let wB = 0, sumB = 0, best = 0, t = 127; const tot = vals.length;
    for (let i = 0; i < bins; i++) {
      wB += hist[i]; if (!wB) continue; const wF = tot - wB; if (!wF) break;
      sumB += i * hist[i]; const mB = sumB / wB, mF = (sum - sumB) / wF, v = wB * wF * (mB - mF) * (mB - mF);
      if (v > best) { best = v; t = i; }
    }
    return lo + t / 255 * span;
  }
  function boxFilter(src, w, h, r) {        // mean over a (2r+1)² box, separable
    const tmp = new Float32Array(w * h), out = new Float32Array(w * h), n = 2 * r + 1;
    for (let y = 0; y < h; y++) {
      let acc = 0; for (let x = -r; x <= r; x++) acc += src[y * w + Math.min(w - 1, Math.max(0, x))];
      for (let x = 0; x < w; x++) {
        tmp[y * w + x] = acc / n;
        acc += src[y * w + Math.min(w - 1, x + r + 1)] - src[y * w + Math.max(0, x - r)];
      }
    }
    for (let x = 0; x < w; x++) {
      let acc = 0; for (let y = -r; y <= r; y++) acc += tmp[Math.min(h - 1, Math.max(0, y)) * w + x];
      for (let y = 0; y < h; y++) {
        out[y * w + x] = acc / n;
        acc += tmp[Math.min(h - 1, y + r + 1) * w + x] - tmp[Math.max(0, y - r) * w + x];
      }
    }
    return out;
  }
  function largestComponent(mask, w, h) {   // 4-connected; returns { count, pts:[x,y,...], touches }
    const seen = new Uint8Array(w * h), stack = new Int32Array(w * h);
    let best = null;
    for (let i = 0; i < w * h; i++) {
      if (!mask[i] || seen[i]) continue;
      let sp = 0, count = 0; stack[sp++] = i; seen[i] = 1; const pts = []; let touches = 0;
      while (sp) {
        const p = stack[--sp], x = p % w, y = (p / w) | 0; count++; pts.push(x, y);
        if (x === 0 || y === 0 || x === w - 1 || y === h - 1) touches++;
        if (x > 0 && mask[p - 1] && !seen[p - 1]) { seen[p - 1] = 1; stack[sp++] = p - 1; }
        if (x < w - 1 && mask[p + 1] && !seen[p + 1]) { seen[p + 1] = 1; stack[sp++] = p + 1; }
        if (y > 0 && mask[p - w] && !seen[p - w]) { seen[p - w] = 1; stack[sp++] = p - w; }
        if (y < h - 1 && mask[p + w] && !seen[p + w]) { seen[p + w] = 1; stack[sp++] = p + w; }
      }
      if (!best || count > best.count) best = { count, pts, touches };
    }
    return best;
  }
  function hull(pts) {                       // monotone chain on flat [x,y,...]; returns [[x,y],...]
    const P = []; for (let i = 0; i < pts.length; i += 2) P.push([pts[i], pts[i + 1]]);
    P.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
    const cross = (o, a, b) => (a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0]);
    const lo = [], up = [];
    for (const p of P) { while (lo.length >= 2 && cross(lo[lo.length - 2], lo[lo.length - 1], p) <= 0) lo.pop(); lo.push(p); }
    for (let i = P.length - 1; i >= 0; i--) { const p = P[i]; while (up.length >= 2 && cross(up[up.length - 2], up[up.length - 1], p) <= 0) up.pop(); up.push(p); }
    up.pop(); lo.pop(); return lo.concat(up);
  }
  function polyArea(q) { let a = 0; for (let i = 0; i < q.length; i++) { const p = q[i], r = q[(i + 1) % q.length]; a += p[0] * r[1] - r[0] * p[1]; } return Math.abs(a) / 2; }
  function minAreaRect(H) {                  // rotating edges of the hull → 4 corners
    let best = null;
    for (let i = 0; i < H.length; i++) {
      const a = H[i], b = H[(i + 1) % H.length]; let ux = b[0] - a[0], uy = b[1] - a[1]; const L = Math.hypot(ux, uy); if (!L) continue;
      ux /= L; uy /= L; const vx = -uy, vy = ux;
      let mnU = Infinity, mxU = -Infinity, mnV = Infinity, mxV = -Infinity;
      for (const p of H) { const u = p[0] * ux + p[1] * uy, v = p[0] * vx + p[1] * vy; mnU = Math.min(mnU, u); mxU = Math.max(mxU, u); mnV = Math.min(mnV, v); mxV = Math.max(mxV, v); }
      const area = (mxU - mnU) * (mxV - mnV);
      if (!best || area < best.area) best = { area, ux, uy, vx, vy, mnU, mxU, mnV, mxV };
    }
    const r = best, P = (u, v) => [u * r.ux + v * r.vx, u * r.uy + v * r.vy];
    return { area: r.area, corners: [P(r.mnU, r.mnV), P(r.mxU, r.mnV), P(r.mxU, r.mxV), P(r.mnU, r.mxV)] };
  }
  function orderQuad(q) {                    // → TL, TR, BR, BL
    const c = q.reduce((a, p) => [a[0] + p[0] / 4, a[1] + p[1] / 4], [0, 0]);
    const s = q.slice().sort((a, b) => Math.atan2(a[1] - c[1], a[0] - c[0]) - Math.atan2(b[1] - c[1], b[0] - c[0]));
    // s is ordered by angle from -π: start from the point with smallest x+y
    let k = 0; for (let i = 1; i < 4; i++) if (s[i][0] + s[i][1] < s[k][0] + s[k][1]) k = i;
    return [0, 1, 2, 3].map((i) => s[(k + i) % 4]);
  }
  function grow(q, f, w, h) {                // push corners outward from the centre by f, clamp to image
    const c = q.reduce((a, p) => [a[0] + p[0] / 4, a[1] + p[1] / 4], [0, 0]);
    return q.map((p) => [Math.max(0, Math.min(w, c[0] + (p[0] - c[0]) * (1 + f))), Math.max(0, Math.min(h, c[1] + (p[1] - c[1]) * (1 + f)))]);
  }

  // ---------- 1. paper vs background ----------
  function byPaper(S) {
    const { w, h, data } = S, n = w * h, paper = new Float32Array(n);
    for (let i = 0; i < n; i++) {
      const r = data[i * 4], g = data[i * 4 + 1], b = data[i * 4 + 2];
      const mx = Math.max(r, g, b), mn = Math.min(r, g, b);
      paper[i] = (r + g + b) / 3 - 1.5 * (mx - mn);           // bright and colourless = paper
    }
    const sm = boxFilter(paper, w, h, 2);
    const t = otsu(sm, -100, 255);
    let m = new Uint8Array(n); for (let i = 0; i < n; i++) m[i] = sm[i] > t ? 1 : 0;
    // close small holes (text) and drop specks: dilate then erode with a 7px box
    const d = boxFilter(Float32Array.from(m), w, h, 3); for (let i = 0; i < n; i++) m[i] = d[i] > 0.02 ? 1 : 0;
    const e = boxFilter(Float32Array.from(m), w, h, 3); for (let i = 0; i < n; i++) m[i] = e[i] > 0.98 ? 1 : 0;
    const comp = largestComponent(m, w, h); if (!comp) return null;
    const frac = comp.count / n;
    const H = hull(comp.pts); if (H.length < 4) return null;
    const rect = minAreaRect(H);
    const fill = comp.count / Math.max(1, rect.area);
    const perim = 2 * (w + h), edgeShare = comp.touches / perim;
    // quad = hull point nearest each rectangle corner (follows perspective a little)
    const quad = rect.corners.map((c) => H.reduce((b, p) => (Math.hypot(p[0] - c[0], p[1] - c[1]) < Math.hypot(b[0] - c[0], b[1] - c[1]) ? p : b), H[0]));
    const qa = polyArea(quad) / n;
    const confident = frac > 0.06 && frac < 0.9 && fill > 0.82 && edgeShare < 0.45 && qa > 0.05;
    return { quad: orderQuad(quad), confident, method: "paper", score: fill };
  }

  // ---------- 2. the two long edges (white on white) ----------
  function byEdges(S) {
    const { w, h, data } = S, n = w * h, g = new Float32Array(n);
    for (let i = 0; i < n; i++) g[i] = 0.299 * data[i * 4] + 0.587 * data[i * 4 + 1] + 0.114 * data[i * 4 + 2];
    const b = boxFilter(g, w, h, 1), mag = new Float32Array(n), gxA = new Float32Array(n), gyA = new Float32Array(n);
    for (let y = 1; y < h - 1; y++) for (let x = 1; x < w - 1; x++) {
      const i = y * w + x;
      const gx = b[i - w + 1] + 2 * b[i + 1] + b[i + w + 1] - b[i - w - 1] - 2 * b[i - 1] - b[i + w - 1];
      const gy = b[i + w - 1] + 2 * b[i + w] + b[i + w + 1] - b[i - w - 1] - 2 * b[i - w] - b[i - w + 1];
      mag[i] = Math.hypot(gx, gy); gxA[i] = gx; gyA[i] = gy;
    }
    const sorted = Float32Array.from(mag).sort(); const T = Math.max(24, sorted[Math.floor(n * 0.9)]);
    // Hough over edge pixels, voting only along the pixel's own gradient direction (±4°) — cheap and clean
    const NT = 180, diag = Math.ceil(Math.hypot(w, h)), NR = 2 * diag + 1, acc = new Float32Array(NT * NR);
    const cosT = new Float32Array(NT), sinT = new Float32Array(NT);
    for (let t = 0; t < NT; t++) { cosT[t] = Math.cos(t * Math.PI / NT); sinT[t] = Math.sin(t * Math.PI / NT); }
    const edges = [];
    for (let y = 2; y < h - 2; y++) for (let x = 2; x < w - 2; x++) {
      const i = y * w + x; if (mag[i] < T) continue;
      edges.push(x, y);
      let th = Math.atan2(gyA[i], gxA[i]); if (th < 0) th += Math.PI; if (th >= Math.PI) th -= Math.PI;
      const t0 = Math.round(th / Math.PI * NT);
      for (let dt = -4; dt <= 4; dt++) { const t = (t0 + dt + NT) % NT; const r = Math.round(x * cosT[t] + y * sinT[t]) + diag; acc[t * NR + r] += 1; }
    }
    // peaks
    const peaks = [];
    for (let k = 0; k < 24; k++) {
      let bi = -1, bv = 0; for (let i = 0; i < acc.length; i++) if (acc[i] > bv) { bv = acc[i]; bi = i; }
      if (bi < 0 || bv < Math.max(w, h) * 0.25) break;
      const t = Math.floor(bi / NR), r = bi % NR - diag; peaks.push({ t, r, v: bv });
      for (let dt = -6; dt <= 6; dt++) for (let dr = -8; dr <= 8; dr++) { const tt = (t + dt + NT) % NT, rr = (dt + t < 0 || dt + t >= NT ? -r : r) + dr + diag; if (rr >= 0 && rr < NR) acc[tt * NR + rr] = 0; }
    }
    // best near-parallel pair, far enough apart, with the most print between them (a receipt is full of text;
    // a desk edge or a mat isn't). Score = weaker line's votes × edge density between the lines.
    const meanDensity = edges.length / 2 / n;
    let pair = null;
    for (let i = 0; i < peaks.length; i++) for (let j = i + 1; j < peaks.length; j++) {
      const a = peaks[i], c = peaks[j]; let dt = Math.abs(a.t - c.t); dt = Math.min(dt, NT - dt); if (dt > 4) continue;
      const rc = Math.abs(a.t - c.t) > NT / 2 ? -c.r : c.r, gap = Math.abs(a.r - rc);
      if (gap < Math.min(w, h) * 0.12) continue;
      const th0 = a.t * Math.PI / NT, cx = Math.cos(th0), cy = Math.sin(th0), lo = Math.min(a.r, rc) + 3, hi = Math.max(a.r, rc) - 3;
      let inside = 0; for (let k = 0; k < edges.length; k += 2) { const r = edges[k] * cx + edges[k + 1] * cy; if (r > lo && r < hi) inside++; }
      // area of the band inside the image (approx: band width × length of the line across the image)
      const len = Math.abs(cx) > Math.abs(cy) ? h / Math.abs(cx) : w / Math.abs(cy);
      const dens = inside / Math.max(1, (hi - lo) * Math.min(len, Math.hypot(w, h)));
      const sc = Math.min(a.v, c.v) * Math.min(3, dens / Math.max(1e-6, meanDensity));
      if (!pair || sc > pair.sc) pair = { a, c: { t: a.t, r: rc }, sc, gap, dens };
    }
    S.lineTheta = peaks.length ? peaks[0].t * Math.PI / NT : null;
    if (!pair) return null;
    const th = pair.a.t * Math.PI / NT, nx = Math.cos(th), ny = Math.sin(th), dx = -ny, dy = nx;  // normal, direction
    const r1 = Math.min(pair.a.r, pair.c.r), r2 = Math.max(pair.a.r, pair.c.r);
    // extent along the lines: where edge pixels (text, ends of the paper) sit between them
    const along = [];
    for (let k = 0; k < edges.length; k += 2) { const x = edges[k], y = edges[k + 1], r = x * nx + y * ny; if (r > r1 + 3 && r < r2 - 3) along.push(x * dx + y * dy); }
    if (along.length < 50) return null;
    along.sort((p, q) => p - q);
    let s1 = along[Math.floor(along.length * 0.01)], s2 = along[Math.floor(along.length * 0.99)];
    const padS = (s2 - s1) * 0.03 + 4; s1 -= padS; s2 += padS;
    const P = (s, r) => [s * dx + r * nx, s * dy + r * ny];
    const clamp = (p) => [Math.max(0, Math.min(w, p[0])), Math.max(0, Math.min(h, p[1]))];
    const quad = [P(s1, r1), P(s2, r1), P(s2, r2), P(s1, r2)].map(clamp);
    const qa = polyArea(quad) / n;
    return { quad: orderQuad(quad), confident: qa > 0.05 && qa < 0.92, method: "edges", score: pair.sc };
  }

  // ---------- 3. where the print is (white receipt on a white desk, edges too faint) ----------
  // Dark strokes on light paper stand out in a "black-hat" map (local max − pixel). The receipt is the band
  // across the photo with the most print; its ends are where the print stops.
  function byPrint(S) {
    const { w, h, data } = S, n = w * h, g = new Float32Array(n);
    for (let i = 0; i < n; i++) g[i] = 0.299 * data[i * 4] + 0.587 * data[i * 4 + 1] + 0.114 * data[i * 4 + 2];
    const r = 3, tmp = new Float32Array(n), mx = new Float32Array(n);
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) { let m = 0; for (let k = -r; k <= r; k++) { const xx = Math.min(w - 1, Math.max(0, x + k)); m = Math.max(m, g[y * w + xx]); } tmp[y * w + x] = m; }
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) { let m = 0; for (let k = -r; k <= r; k++) { const yy = Math.min(h - 1, Math.max(0, y + k)); m = Math.max(m, tmp[yy * w + x]); } mx[y * w + x] = m; }
    const bh = new Float32Array(n); for (let i = 0; i < n; i++) bh[i] = mx[i] - g[i];
    const srt = Float32Array.from(bh).sort(); const T = Math.max(25, srt[Math.floor(n * 0.94)]);
    const pts = []; for (let y = 3; y < h - 3; y++) for (let x = 3; x < w - 3; x++) if (bh[y * w + x] > T) pts.push(x, y);
    if (pts.length < 200) return null;
    const thetas = [];
    if (S.lineTheta != null) thetas.push(S.lineTheta);
    thetas.push(Math.PI / 2, 0);
    let best = null;
    for (const th of thetas) {
      const nx = Math.cos(th), ny = Math.sin(th), dx = -ny, dy = nx, diag = Math.ceil(Math.hypot(w, h)), hist = new Float32Array(2 * diag + 1);
      for (let k = 0; k < pts.length; k += 2) hist[Math.round(pts[k] * nx + pts[k + 1] * ny) + diag]++;
      // smooth over 9px so single lines blur out and text rows merge into a band
      const sm = new Float32Array(hist.length); for (let i = 0; i < hist.length; i++) { let a = 0; for (let k = -6; k <= 6; k++) a += hist[Math.min(hist.length - 1, Math.max(0, i + k))]; sm[i] = a / 13; }
      let peak = 0; for (const v of sm) peak = Math.max(peak, v);
      const cut = peak * 0.08;
      // runs above the cut; keep the one holding most print
      let run = null, i0 = -1, mass = 0;
      for (let i = 0; i <= sm.length; i++) {
        if (i < sm.length && sm[i] > cut) { if (i0 < 0) { i0 = i; mass = 0; } mass += hist[i]; }
        else if (i0 >= 0) { if (!run || mass > run.mass) run = { a: i0 - diag, b: i - 1 - diag, mass }; i0 = -1; }
      }
      if (!run) continue;
      const width = run.b - run.a, share = run.mass / (pts.length / 2);
      if (width < Math.min(w, h) * 0.1) continue;
      const along = []; for (let k = 0; k < pts.length; k += 2) { const rr = pts[k] * nx + pts[k + 1] * ny; if (rr >= run.a && rr <= run.b) along.push(pts[k] * dx + pts[k + 1] * dy); }
      along.sort((p, q) => p - q);
      const s1 = along[Math.floor(along.length * 0.01)], s2 = along[Math.floor(along.length * 0.99)];
      // concentration: how much more print is in the band than if print were spread evenly across the photo
      const span = Math.abs(nx) * w + Math.abs(ny) * h, conc = share / Math.max(0.05, width / span);
      const score = conc * Math.min(1, (s2 - s1) / Math.max(w, h) + 0.3);
      if (!best || score > best.score) best = { th, nx, ny, dx, dy, a: run.a, b: run.b, s1, s2, score, share, conc };
    }
    if (!best || best.share < 0.3 || best.conc < 1.5) return null;
    const padR = (best.b - best.a) * 0.03 + 3, padS = (best.s2 - best.s1) * 0.03 + 4;
    const r1 = best.a - padR, r2 = best.b + padR, s1 = best.s1 - padS, s2 = best.s2 + padS;
    const P = (sv, rv) => [sv * best.dx + rv * best.nx, sv * best.dy + rv * best.ny];
    const clamp = (p) => [Math.max(0, Math.min(w, p[0])), Math.max(0, Math.min(h, p[1]))];
    const quad = [P(s1, r1), P(s2, r1), P(s2, r2), P(s1, r2)].map(clamp);
    const qa = polyArea(quad) / n;
    return { quad: orderQuad(quad), confident: qa > 0.05 && qa < 0.85, method: "print", score: best.score };
  }

  function analyse(img) {
    const S = smallCopy(img);
    const iw = img.naturalWidth || img.width, ih = img.naturalHeight || img.height;
    let r = null;
    try { r = byPaper(S); } catch (_) { r = null; }
    if (!r || !r.confident) { let e = null; try { e = byEdges(S); } catch (_) { e = null; } if (e && e.confident) r = e; }
    if (!r || !r.confident) { let q = null; try { q = byPrint(S); } catch (_) { q = null; } if (q && q.confident) r = q; }
    let quad, method = "whole", confident = false;
    if (r && r.confident) { quad = grow(r.quad, 0.01, S.w, S.h).map((p) => [p[0] / S.s, p[1] / S.s]); method = r.method; confident = true; }
    else quad = [[0, 0], [iw, 0], [iw, ih], [0, ih]];
    // a receipt lying sideways: turn it upright (the person can turn it again)
    const wq = (Math.hypot(quad[1][0] - quad[0][0], quad[1][1] - quad[0][1]) + Math.hypot(quad[2][0] - quad[3][0], quad[2][1] - quad[3][1])) / 2;
    const hq = (Math.hypot(quad[3][0] - quad[0][0], quad[3][1] - quad[0][1]) + Math.hypot(quad[2][0] - quad[1][0], quad[2][1] - quad[1][1])) / 2;
    const rot = confident && wq > hq * 1.25 ? 270 : 0;
    return { quad, rot, method, confident };
  }

  // ---------- straighten (perspective) ----------
  function homography(src, dst) {            // maps dst → src (for inverse sampling)
    const A = [], B = [];
    for (let i = 0; i < 4; i++) {
      const [x, y] = dst[i], [u, v] = src[i];
      A.push([x, y, 1, 0, 0, 0, -u * x, -u * y]); B.push(u);
      A.push([0, 0, 0, x, y, 1, -v * x, -v * y]); B.push(v);
    }
    for (let c = 0; c < 8; c++) {             // Gaussian elimination
      let p = c; for (let r = c + 1; r < 8; r++) if (Math.abs(A[r][c]) > Math.abs(A[p][c])) p = r;
      [A[c], A[p]] = [A[p], A[c]]; [B[c], B[p]] = [B[p], B[c]];
      for (let r = 0; r < 8; r++) { if (r === c) continue; const f = A[r][c] / A[c][c]; for (let k = c; k < 8; k++) A[r][k] -= f * A[c][k]; B[r] -= f * B[c]; }
    }
    return B.map((b, i) => b / A[i][i]).concat([1]);
  }
  function warp(img, quad, maxEdge) {
    const d = (a, b) => Math.hypot(a[0] - b[0], a[1] - b[1]);
    let W = Math.max(d(quad[0], quad[1]), d(quad[3], quad[2])), H = Math.max(d(quad[0], quad[3]), d(quad[1], quad[2]));
    const s = Math.min(1.5, maxEdge / Math.max(W, H)); W = Math.max(16, Math.round(W * s)); H = Math.max(16, Math.round(H * s));
    const iw = img.naturalWidth || img.width, ih = img.naturalHeight || img.height;
    const sc = document.createElement("canvas"); sc.width = iw; sc.height = ih;
    const sx = sc.getContext("2d", { willReadFrequently: true }); sx.drawImage(img, 0, 0);
    const src = sx.getImageData(0, 0, iw, ih).data;
    const out = document.createElement("canvas"); out.width = W; out.height = H;
    const ox = out.getContext("2d"), od = ox.createImageData(W, H), o = od.data;
    const M = homography(quad, [[0, 0], [W, 0], [W, H], [0, H]]);
    for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
      const z = M[6] * x + M[7] * y + 1, u = (M[0] * x + M[1] * y + M[2]) / z, v = (M[3] * x + M[4] * y + M[5]) / z;
      const x0 = Math.max(0, Math.min(iw - 2, Math.floor(u))), y0 = Math.max(0, Math.min(ih - 2, Math.floor(v)));
      const fx = Math.max(0, Math.min(1, u - x0)), fy = Math.max(0, Math.min(1, v - y0));
      const i00 = (y0 * iw + x0) * 4, i10 = i00 + 4, i01 = i00 + iw * 4, i11 = i01 + 4, j = (y * W + x) * 4;
      for (let c = 0; c < 3; c++) o[j + c] = (src[i00 + c] * (1 - fx) + src[i10 + c] * fx) * (1 - fy) + (src[i01 + c] * (1 - fx) + src[i11 + c] * fx) * fy;
      o[j + 3] = 255;
    }
    ox.putImageData(od, 0, 0); return out;
  }
  function rotate(c, deg) {
    if (!deg) return c;
    const o = document.createElement("canvas"), q = deg % 180 !== 0;
    o.width = q ? c.height : c.width; o.height = q ? c.width : c.height;
    const x = o.getContext("2d"); x.translate(o.width / 2, o.height / 2); x.rotate(deg * Math.PI / 180); x.drawImage(c, -c.width / 2, -c.height / 2);
    return o;
  }

  // ---------- make the text stand out ----------
  function enhance(c) {
    const w = c.width, h = c.height, x = c.getContext("2d", { willReadFrequently: true }), id = x.getImageData(0, 0, w, h), d = id.data, n = w * h;
    // background = a heavily blurred copy (paper colour + shadows), made by shrinking and growing again
    const k = Math.max(8, Math.round(Math.max(w, h) / 40));
    const sm = document.createElement("canvas"); sm.width = Math.max(4, Math.round(w / k)); sm.height = Math.max(4, Math.round(h / k));
    const smx = sm.getContext("2d", { willReadFrequently: true }); smx.imageSmoothingQuality = "high"; smx.drawImage(c, 0, 0, sm.width, sm.height);
    // take the brightest of each neighbourhood so dark text doesn't count as background
    const sd = smx.getImageData(0, 0, sm.width, sm.height), sw = sm.width, sh = sm.height, g0 = new Float32Array(sw * sh);
    for (let i = 0; i < sw * sh; i++) g0[i] = 0.299 * sd.data[i * 4] + 0.587 * sd.data[i * 4 + 1] + 0.114 * sd.data[i * 4 + 2];
    const g1 = new Float32Array(sw * sh);
    for (let y = 0; y < sh; y++) for (let xx = 0; xx < sw; xx++) {
      let m = 0; for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) { const yy = Math.min(sh - 1, Math.max(0, y + dy)), x2 = Math.min(sw - 1, Math.max(0, xx + dx)); m = Math.max(m, g0[yy * sw + x2]); }
      g1[y * sw + xx] = m;
    }
    for (let i = 0; i < sw * sh; i++) { const v = g1[i]; sd.data[i * 4] = sd.data[i * 4 + 1] = sd.data[i * 4 + 2] = v; }
    smx.putImageData(sd, 0, 0);
    const bg = document.createElement("canvas"); bg.width = w; bg.height = h;
    const bgx = bg.getContext("2d", { willReadFrequently: true }); bgx.imageSmoothingEnabled = true; bgx.imageSmoothingQuality = "high"; bgx.filter = "blur(" + Math.round(k / 2) + "px)";
    bgx.drawImage(sm, 0, 0, w, h);
    const b = bgx.getImageData(0, 0, w, h).data;
    const g = new Float32Array(n), hist = new Uint32Array(256);
    for (let i = 0; i < n; i++) {
      const y = 0.299 * d[i * 4] + 0.587 * d[i * 4 + 1] + 0.114 * d[i * 4 + 2], bb = Math.max(40, b[i * 4]);
      const v = Math.min(255, y / bb * 255); g[i] = v; hist[v | 0]++;
    }
    // black point = darkest 1%; paper → white; gentle curve so faint thermal print gets darker
    let acc = 0, lo = 0; for (let i = 0; i < 256; i++) { acc += hist[i]; if (acc > n * 0.01) { lo = i; break; } }
    lo = Math.min(lo, 120); const hi = 235;
    for (let i = 0; i < n; i++) {
      let v = (g[i] - lo) / (hi - lo); v = Math.max(0, Math.min(1, v)); v = Math.pow(v, 1.6) * 255;
      d[i * 4] = d[i * 4 + 1] = d[i * 4 + 2] = v;
    }
    x.putImageData(id, 0, 0); return c;
  }

  async function render(img, opt = {}) {
    const { quad, rot = 0, enhance: enh = true, maxEdge = 1600, quality = 0.82 } = opt;
    let c = warp(img, quad, maxEdge);
    c = rotate(c, rot);
    if (enh) enhance(c);
    return new Promise((res) => c.toBlob(res, "image/jpeg", quality));
  }

  window.MuirScan = { analyse, render, warp, enhance, rotate, _byPaper: byPaper, _byEdges: byEdges, _byPrint: byPrint, _small: smallCopy };
})();
