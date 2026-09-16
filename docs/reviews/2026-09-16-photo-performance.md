# Photo display measurements (2026-09-16)

## Scope and method

- Target: the owner-only development `/dev` deployment in `gas-project.json`.
- Browser: the same signed-in Chrome session on Windows, with real GAS, Sheets,
  and Drive requests. No artificial network delay or media-response stub.
- Fixture: one existing 1920 x 1440 photo, unchanged throughout the comparison.
  No photo contents, file IDs, edit keys, or authenticated URLs are recorded here.
- Start: immediately before `openPinDetail` for that photo. End: the production
  photo loader resolves, `HTMLImageElement.decode()` resolves, and two animation
  frames pass. Close the detail between samples. This includes decoding and a
  paint opportunity but is not a measurement of a human observer's reaction time.
- Keep first fetches separate from reuse within the same page. Navigation clears
  the application cache. GAS latency is variable; these observations do not
  establish a universal speedup or an improvement to initial application startup.
- The test data contains no audio-enabled pins, so no real-GAS music speedup is
  claimed. Existing audio caching and playback behavior are unchanged.

## Preliminary measurements (not used for the acceptance decision)

Before the change, consecutive openings of the same photo took:

| Sample | Display time (ms) |
| --- | ---: |
| 1 | 5892 |
| 2 | 6422 |
| 3 | 6005 |
| 4 | 6841 |
| 5 | 6656 |
| 6 | 6222 |

The first sample was the first opening in this page. The five subsequent
openings had a median of 6422 ms (range 6005-6841 ms). The old loader discarded
its only Object URL on close and requested the bytes again on every reopening.

An initial post-change run returned 5261 ms for the first fetch and 2001 ms for
each of five subsequent views. Follow-up inspection found the tab was not
focused. After bringing Chrome to the foreground, a warm sample spent 3.9 ms
waiting for the loader, 37.7 ms through decoding, and 58.3 ms through two frames.
The earlier display timings include browser frame throttling and are not used
to quantify the improvement.

## Controlled comparison and decision

The following comparison used the candidate code in the foreground, alternating
cache misses and hits. Before each miss, close the detail and clear only
`pinPhotoCache`. Immediately reopen normally for the hit. Both cases use the
same real authenticated fetch path and identical rendering code; this isolates
the effect of reusing photo bytes. This is an on/off experiment within the new
build, not a claim that the old deployment was re-measured in the foreground.
All ten samples recorded a focused, visible document and a decoded 1920 x 1440
image.

| Pair | Miss: loader (ms) | Miss: decoded (ms) | Miss: displayed (ms) | Hit: loader (ms) | Hit: decoded (ms) | Hit: displayed (ms) |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| 1 | 4776.9 | 4802.3 | 4828.7 | 3.2 | 26.8 | 49.2 |
| 2 | 4264.4 | 4287.2 | 4303.9 | 2.5 | 26.8 | 49.5 |
| 3 | 4666.1 | 4694.6 | 4720.9 | 2.3 | 25.9 | 49.6 |
| 4 | 4466.1 | 4491.2 | 4514.7 | 2.3 | 25.5 | 48.1 |
| 5 | 6253.6 | 6281.5 | 6306.5 | 2.3 | 25.9 | 48.2 |

Median displayed time: 4720.9 ms without a cached entry, 49.2 ms with one.
Ranges: 4303.9-6306.5 ms and 48.1-49.6 ms respectively. Adopt the bounded photo
cache for repeated views. This does not improve the first fetch, and is not
evidence of a music, shared-view, mobile-network, or startup speedup.

The on/off probe can be repeated in the development edit page's application
iframe, after initialization and with the browser foregrounded:

```js
const measurements = [];
const photo = state.pins.find(pin => pin.fileId);
const image = document.getElementById('pin-detail-image');
for (let pair = 1; pair <= 5; pair += 1) {
  for (const mode of ['uncached', 'cached']) {
    closePinDetail({ restoreFocus: false });
    if (mode === 'uncached') pinPhotoCache.clear();
    if (!document.hasFocus() || document.visibilityState !== 'visible') {
      throw new Error('Foreground the browser before measuring.');
    }
    const start = performance.now();
    openPinDetail(photo);
    if (!await pinPhotoLoader.open(photo.id, photo.title)) throw new Error('Photo load failed.');
    const readyMs = performance.now() - start;
    await image.decode();
    const decodedMs = performance.now() - start;
    await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
    measurements.push({ pair, mode, readyMs, decodedMs, displayMs: performance.now() - start });
  }
}
console.table(measurements);
```

## Verification

- `validate`: 1584 Node tests and 55 browser tests passed, including cache bounds,
  expiry, revision changes, stale responses, retry, map/detail reuse, and existing
  photo/audio behavior.
- Reviewed the remote snapshot against the baseline before pushing. Test GAS
  HEAD was updated via `gas:push` and read back with matching source files.
- Actual `/dev` photo detail and map popup were visually checked. Existing photo
  files and pin data were not edited. The existing sheet menu refreshed its edit
  URL using the existing key; no key regeneration or sharing change was made.

## Candidate behavior

The edit page shares validated photo Blobs between the map popup and pin detail,
up to three entries and 24 MiB total. Each entry expires 60 seconds after it was
stored; revisiting does not extend this time. Keys include pin ID, photo source,
and the known update timestamp. Reopening after a known revision fetches again.
Retry removes the cached entry. Invalid, failed, closed, or obsolete responses
are not cached. Object URLs still belong to each view and are revoked on close
or destruction. The page-wide cache is cleared on `pagehide`.

Another client can change the underlying file before this page learns of that
revision; a cached photo can remain visible on reopening for up to 60 seconds.
The cache is memory-only and does not add persistent browser storage or change
server authentication, Drive permissions, sheet columns, or the fixed `/exec`
deployment.
