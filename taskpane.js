/* ============================================================
   Character Remover — Task Pane Logic
   Works in Microsoft Word and Microsoft Excel.
   ============================================================ */

"use strict";

/* ── State ─────────────────────────────────────────────── */
/** @type {Set<string>} Each entry is a single Unicode code-point string */
const charSet = new Set();

/* ── DOM refs ───────────────────────────────────────────── */
const charInput   = document.getElementById("charInput");
const btnAdd      = document.getElementById("btnAdd");
const tagList     = document.getElementById("tagList");
const btnRemove   = document.getElementById("btnRemove");
const btnClearAll = document.getElementById("btnClearAll");
const statusBox   = document.getElementById("statusBox");
const optAllSheets = document.getElementById("optAllSheets");

/* ── Office initialisation ──────────────────────────────── */
Office.onReady(function (info) {
  // Hide the "all sheets" option when running inside Word
  if (info.host === Office.HostType.Word) {
    document.getElementById("optAllSheets").closest(".option-row").style.display = "none";
  }
});

/* ── Helpers ────────────────────────────────────────────── */

/**
 * Split a string into individual Unicode grapheme clusters so that
 * multi-codepoint emoji (e.g. 👨‍👩‍👧) are treated as one character.
 * Falls back to codepoint-by-codepoint splitting when the Intl
 * Segmenter API is unavailable (older Office WebView versions).
 *
 * @param {string} str
 * @returns {string[]}
 */
function splitGraphemes(str) {
  if (typeof Intl !== "undefined" && typeof Intl.Segmenter === "function") {
    const segmenter = new Intl.Segmenter();
    return Array.from(segmenter.segment(str), s => s.segment);
  }
  // Fallback: split by Unicode code points
  return Array.from(str);
}

/**
 * Show a status message in the banner.
 * @param {string} message
 * @param {"success"|"error"|"info"} type
 */
function showStatus(message, type) {
  statusBox.textContent = message;
  statusBox.className = "status show " + type;
}

function hideStatus() {
  statusBox.className = "status";
}

/**
 * Rebuild the tag list UI from the current charSet.
 */
function renderTags() {
  tagList.innerHTML = "";

  if (charSet.size === 0) {
    tagList.classList.add("empty");
    btnRemove.disabled = true;
    return;
  }

  tagList.classList.remove("empty");
  btnRemove.disabled = false;

  charSet.forEach(function (ch) {
    const tag = document.createElement("span");
    tag.className = "tag";
    tag.title = describeChar(ch);

    // Character display
    const charSpan = document.createElement("span");
    charSpan.textContent = ch;
    tag.appendChild(charSpan);

    // Remove button
    const removeBtn = document.createElement("button");
    removeBtn.className = "tag-remove";
    removeBtn.innerHTML = "&#x2715;"; // ×
    removeBtn.title = "Remove \"" + ch + "\" from list";
    removeBtn.setAttribute("aria-label", "Remove character " + ch);
    removeBtn.addEventListener("click", function () {
      charSet.delete(ch);
      renderTags();
      hideStatus();
    });
    tag.appendChild(removeBtn);

    tagList.appendChild(tag);
  });
}

/**
 * Return a human-readable description of a character for tooltip use.
 * @param {string} ch
 * @returns {string}
 */
function describeChar(ch) {
  const codePoints = Array.from(ch)
    .map(c => "U+" + c.codePointAt(0).toString(16).toUpperCase().padStart(4, "0"))
    .join(" ");
  return ch + "  (" + codePoints + ")";
}

/**
 * Add all grapheme clusters from the input field to charSet.
 */
function addCharsFromInput() {
  const raw = charInput.value.trim();
  if (!raw) return;

  const clusters = splitGraphemes(raw);
  let added = 0;
  clusters.forEach(function (cluster) {
    if (!charSet.has(cluster)) {
      charSet.add(cluster);
      added++;
    }
  });

  charInput.value = "";
  charInput.focus();
  renderTags();

  if (added > 0) {
    hideStatus();
  }
}

/* ── Event listeners ────────────────────────────────────── */

btnAdd.addEventListener("click", addCharsFromInput);

charInput.addEventListener("keydown", function (e) {
  if (e.key === "Enter") {
    e.preventDefault();
    addCharsFromInput();
  }
});

btnClearAll.addEventListener("click", function () {
  charSet.clear();
  renderTags();
  hideStatus();
});

btnRemove.addEventListener("click", removeCharacters);

/* ── Core removal logic ─────────────────────────────────── */

/**
 * Entry point: detect the Office host and dispatch to the
 * appropriate removal function.
 */
async function removeCharacters() {
  if (charSet.size === 0) {
    showStatus("Please add at least one character before removing.", "error");
    return;
  }

  setLoading(true);
  hideStatus();

  try {
    const host = Office.context.host;

    if (host === Office.HostType.Word) {
      await removeFromWord();
    } else if (host === Office.HostType.Excel) {
      await removeFromExcel();
    } else {
      showStatus("Unsupported Office host. This add-in works in Word and Excel only.", "error");
    }
  } catch (err) {
    console.error("Character Remover error:", err);
    showStatus("An error occurred: " + (err.message || String(err)), "error");
  } finally {
    setLoading(false);
  }
}

/**
 * Toggle the loading state of the Remove button.
 * @param {boolean} loading
 */
function setLoading(loading) {
  if (loading) {
    btnRemove.classList.add("loading");
    btnRemove.disabled = true;
  } else {
    btnRemove.classList.remove("loading");
    btnRemove.disabled = charSet.size === 0;
  }
}

/* ── Word removal ───────────────────────────────────────── */

/**
 * Remove all target characters from the active Word document.
 *
 * Strategy: for each target character, use Word's built-in
 * body.search() to find every occurrence and replace it with
 * an empty string.  This is the most reliable approach because
 * it handles characters inside text runs, headers, footers,
 * footnotes, etc., through the document body search.
 *
 * Note: Word.Body.search() operates on the main body. For
 * headers/footers the same pattern is applied via the sections API.
 */
async function removeFromWord() {
  let totalRemoved = 0;

  await Word.run(async function (context) {
    const chars = Array.from(charSet);

    for (const ch of chars) {
      // Search in the main body
      const results = context.document.body.search(ch, {
        matchCase: true,
        matchWholeWord: false,
        matchWildcards: false,
      });
      context.load(results, "items");
      await context.sync();

      const count = results.items.length;
      totalRemoved += count;

      for (const range of results.items) {
        range.insertText("", Word.InsertLocation.replace);
      }
      await context.sync();
    }

    // Also clean headers and footers on each section
    const sections = context.document.sections;
    context.load(sections, "items");
    await context.sync();

    for (const section of sections.items) {
      const headerFooterTypes = [
        Word.HeaderFooterType.primary,
        Word.HeaderFooterType.firstPage,
        Word.HeaderFooterType.evenPages,
      ];

      for (const hfType of headerFooterTypes) {
        for (const isHeader of [true, false]) {
          try {
            const hf = isHeader
              ? section.getHeader(hfType)
              : section.getFooter(hfType);

            for (const ch of chars) {
              const hfResults = hf.search(ch, {
                matchCase: true,
                matchWholeWord: false,
                matchWildcards: false,
              });
              context.load(hfResults, "items");
              await context.sync();

              totalRemoved += hfResults.items.length;
              for (const range of hfResults.items) {
                range.insertText("", Word.InsertLocation.replace);
              }
              await context.sync();
            }
          } catch (_) {
            // Header/footer may not exist — silently skip
          }
        }
      }
    }
  });

  const charLabel = charSet.size === 1 ? "character" : "characters";
  showStatus(
    "Done! Removed " + totalRemoved + " instance" + (totalRemoved !== 1 ? "s" : "") +
    " of " + charSet.size + " " + charLabel + " from the document.",
    "success"
  );
}

/* ── Excel removal ──────────────────────────────────────── */

/**
 * Remove all target characters from the active Excel workbook.
 *
 * Strategy: use Worksheet.getUsedRange() to get only cells that
 * contain data, then use Range.replaceAll() for each target
 * character.  This is fast even for large workbooks.
 *
 * The "all sheets" option iterates over every worksheet;
 * otherwise only the active sheet is processed.
 */
async function removeFromExcel() {
  let totalRemoved = 0;
  const processAllSheets = optAllSheets.checked;

  await Excel.run(async function (context) {
    const chars = Array.from(charSet);

    /** @type {Excel.Worksheet[]} */
    let sheets;

    if (processAllSheets) {
      const sheetCollection = context.workbook.worksheets;
      sheetCollection.load("items");
      await context.sync();
      sheets = sheetCollection.items;
    } else {
      sheets = [context.workbook.worksheets.getActiveWorksheet()];
      await context.sync();
    }

    for (const sheet of sheets) {
      let usedRange;
      try {
        usedRange = sheet.getUsedRange();
        usedRange.load("address");
        await context.sync();
      } catch (_) {
        // Sheet may be empty — skip
        continue;
      }

      for (const ch of chars) {
        const result = await usedRange.replaceAll(ch, "", {
          completeMatch: false,
          matchCase: true,
        });
        await context.sync();
        totalRemoved += result.value || 0;
      }
    }
  });

  const sheetLabel = processAllSheets ? "all worksheets" : "the active worksheet";
  const charLabel  = charSet.size === 1 ? "character" : "characters";
  showStatus(
    "Done! Removed " + totalRemoved + " instance" + (totalRemoved !== 1 ? "s" : "") +
    " of " + charSet.size + " " + charLabel + " across " + sheetLabel + ".",
    "success"
  );
}
