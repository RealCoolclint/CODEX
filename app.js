(function () {
  const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

  const dropzone = document.getElementById("dropzone");
  const resultsOutput = document.getElementById("results-output");
  const downloadActions = document.getElementById("download-actions");
  const btnDownloadPrompteur = document.getElementById("btn-download-prompteur");
  const btnDownloadFiche = document.getElementById("btn-download-fiche");
  const btnDownloadBoth = document.getElementById("btn-download-both");
  const wpmInput = document.getElementById("wpm");
  let currentModel = null;

  document.addEventListener("dragover", function (event) {
    event.preventDefault();
  });

  document.addEventListener("drop", function (event) {
    event.preventDefault();
  });

  dropzone.addEventListener("dragover", function (event) {
    event.preventDefault();
    if (event.dataTransfer) event.dataTransfer.dropEffect = "copy";
    dropzone.classList.add("drag-over");
  });

  dropzone.addEventListener("dragleave", function (event) {
    event.preventDefault();
    dropzone.classList.remove("drag-over");
  });

  dropzone.addEventListener("drop", function (event) {
    event.preventDefault();
    dropzone.classList.remove("drag-over");
    const file = event.dataTransfer && event.dataTransfer.files[0];
    readDocx(file);
  });

  const fileInput = document.getElementById("file-input");

  fileInput.addEventListener("click", function (event) {
    event.stopPropagation();
  });

  dropzone.addEventListener("click", function () {
    fileInput.click();
  });

  fileInput.addEventListener("change", function () {
    const file = fileInput.files && fileInput.files[0];
    fileInput.value = "";
    readDocx(file);
  });

  wpmInput.addEventListener("input", function () {
    if (currentModel) renderModel(currentModel);
  });

  function attr(el, name) {
    if (!el) return null;
    const named = el.getAttributeNS(W_NS, name);
    if (named != null) return named;
    const prefixed = el.getAttribute("w:" + name);
    if (prefixed != null) return prefixed;
    return el.getAttribute(name);
  }

  function directChild(el, localName) {
    if (!el) return null;
    for (let i = 0; i < el.children.length; i++) {
      const child = el.children[i];
      if (child.localName === localName && child.namespaceURI === W_NS) return child;
    }
    return null;
  }

  function belongsTo(node, root) {
    let parent = node.parentElement;
    while (parent && parent !== root) {
      if (parent.localName === "p" && parent.namespaceURI === W_NS) return false;
      parent = parent.parentElement;
    }
    return parent === root;
  }

  function isBold(rPr) {
    const bold = directChild(rPr, "b");
    if (!bold) return false;
    const val = attr(bold, "val");
    if (val == null || val === "") return true;
    const normalized = val.trim().toLowerCase();
    return normalized !== "0" && normalized !== "false" && normalized !== "off";
  }

  function readSize(rPr) {
    const sz = directChild(rPr, "sz");
    if (!sz) return null;
    const raw = attr(sz, "val");
    if (raw == null || raw.trim() === "") return null;
    const value = Number(raw);
    return Number.isFinite(value) ? value : null;
  }

  function readHighlight(rPr) {
    const highlight = directChild(rPr, "highlight");
    if (!highlight) return null;
    const val = attr(highlight, "val");
    if (!val || val.toLowerCase() === "none") return null;
    return val;
  }

  function readShadeFill(rPr) {
    const shd = directChild(rPr, "shd");
    if (!shd) return null;
    const val = attr(shd, "fill");
    if (val == null || val.trim() === "") return null;
    const normalized = val.replace(/\s+/g, "").toUpperCase();
    if (normalized === "AUTO") return null;
    return normalized;
  }

  function isItalic(rPr) {
    const italic = directChild(rPr, "i");
    if (!italic) return false;
    const val = attr(italic, "val");
    if (val == null || val === "") return true;
    const normalized = val.trim().toLowerCase();
    return normalized !== "0" && normalized !== "false" && normalized !== "off";
  }

  function readColor(rPr) {
    const color = directChild(rPr, "color");
    if (!color) return null;
    const val = attr(color, "val");
    if (val == null || val.trim() === "") return null;
    return val.trim();
  }

  function collectParagraphRuns(paragraph) {
    const runs = [];

    function collectFromElement(element) {
      for (let i = 0; i < element.children.length; i++) {
        const child = element.children[i];
        if (child.namespaceURI !== W_NS) continue;
        if (child.localName === "r") {
          runs.push(child);
        } else if (child.localName === "hyperlink") {
          collectFromElement(child);
        }
      }
    }

    collectFromElement(paragraph);
    return runs;
  }

  function parseParagraph(paragraph) {
    const runNodes = collectParagraphRuns(paragraph);
    const runs = [];
    let paragraphBold = false;
    let paragraphSize = null;
    let representativeFound = false;
    const pPr = directChild(paragraph, "pPr");
    const paragraphShadeFill = pPr ? readShadeFill(pPr) : null;
    const paragraphInfographic = paragraphShadeFill === "BDD7EE";

    for (let i = 0; i < runNodes.length; i++) {
      const run = runNodes[i];

      const rPr = directChild(run, "rPr");
      let text = "";
      const texts = run.getElementsByTagNameNS(W_NS, "t");
      for (let j = 0; j < texts.length; j++) {
        if (!belongsTo(texts[j], paragraph)) continue;
        text += texts[j].textContent || "";
      }

      const runShadeFill = rPr ? readShadeFill(rPr) : null;
      runs.push({
        text: text,
        bold: rPr ? isBold(rPr) : false,
        highlight: rPr ? readHighlight(rPr) : null,
        italic: rPr ? isItalic(rPr) : false,
        color: rPr ? readColor(rPr) : null,
        infographic: runShadeFill === "BDD7EE" || paragraphInfographic
      });

      if (!representativeFound && text.trim() !== "") {
        representativeFound = true;
        paragraphBold = rPr ? isBold(rPr) : false;
        paragraphSize = rPr ? readSize(rPr) : null;
      }
    }

    return {
      paragraphBold: paragraphBold,
      paragraphSize: paragraphSize,
      runs: runs
    };
  }

  function parseDocumentXml(xml) {
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    if (doc.getElementsByTagName("parsererror").length) {
      throw new Error("XML invalide");
    }

    const paragraphs = doc.getElementsByTagNameNS(W_NS, "p");
    const parsed = [];
    for (let i = 0; i < paragraphs.length; i++) {
      parsed.push(parseParagraph(paragraphs[i]));
    }
    return parsed;
  }

  function paragraphText(paragraph) {
    let text = "";
    for (let i = 0; i < paragraph.runs.length; i++) text += paragraph.runs[i].text;
    return text;
  }

  function isMetaParagraph(paragraph) {
    let hasText = false;
    for (let i = 0; i < paragraph.runs.length; i++) {
      const run = paragraph.runs[i];
      if (run.text.trim() === "") continue;
      hasText = true;
      const color = run.color ? run.color.trim().toLowerCase() : "";
      if (!run.italic || color !== "666666") return false;
    }
    return hasText;
  }

  function isInsertParagraph(paragraph) {
    if (paragraph.paragraphBold && paragraph.paragraphSize === 24) return false;
    let hasNonEmpty = false;
    for (let i = 0; i < paragraph.runs.length; i++) {
      const run = paragraph.runs[i];
      if (run.text.trim() === "") continue;
      hasNonEmpty = true;
      if (!run.bold) return false;
    }
    return hasNonEmpty;
  }

  function cameraFromHighlight(highlight) {
    if (!highlight) return 1;
    return highlight.trim().toLowerCase() === "yellow" ? 2 : 1;
  }

  function lineCountForText(text) {
    if (text === "") return 1;
    return String(text).split("\n").length;
  }

  function emptyInfographicsByLine(lineCount) {
    const lines = [];
    for (let i = 0; i < lineCount; i++) lines.push([]);
    return lines;
  }

  function syncInfographicsByLine(block) {
    const needed = lineCountForText(block.text);
    while (block.infographicsByLine.length < needed) block.infographicsByLine.push([]);
    block.infographicsByLine.length = needed;
  }

  function cloneInfographicsByLine(infographicsByLine) {
    return infographicsByLine.map(function (line) {
      return line.slice();
    });
  }

  function lineIndexFromCharOffset(text, charIndex) {
    return text.slice(0, charIndex).split("\n").length - 1;
  }

  function extractInfographicPassages(runs) {
    const passages = [];
    let current = null;

    for (let i = 0; i < runs.length; i++) {
      const text = runs[i].text;
      if (text.trim() === "") {
        if (current) current.text += text;
        continue;
      }
      if (runs[i].infographic) {
        if (current) {
          current.text += text;
        } else {
          current = { startRunIndex: i, text: text };
        }
      } else if (current) {
        passages.push({ startRunIndex: current.startRunIndex, text: current.text.trim() });
        current = null;
      }
    }

    if (current) {
      passages.push({ startRunIndex: current.startRunIndex, text: current.text.trim() });
    }

    return passages;
  }

  function runsToBlocks(runs) {
    const blocks = [];
    const runPositions = [];
    const passages = extractInfographicPassages(runs);
    let leading = "";

    for (let i = 0; i < runs.length; i++) {
      const text = runs[i].text;
      if (text.trim() === "") {
        if (blocks.length) {
          const trailingBlock = blocks[blocks.length - 1];
          trailingBlock.text += text;
          syncInfographicsByLine(trailingBlock);
        } else {
          leading += text;
        }
        continue;
      }

      const camera = cameraFromHighlight(runs[i].highlight);
      const last = blocks[blocks.length - 1];

      if (last && last.type === "camera" && last.camera === camera) {
        runPositions[i] = { blockIndex: blocks.length - 1, charIndex: last.text.length };
        last.text += text;
        syncInfographicsByLine(last);
      } else {
        const blockText = leading + text;
        runPositions[i] = { blockIndex: blocks.length, charIndex: leading.length };
        blocks.push({
          type: "camera",
          camera: camera,
          text: blockText,
          infographicsByLine: emptyInfographicsByLine(lineCountForText(blockText))
        });
        leading = "";
      }
    }

    for (let p = 0; p < passages.length; p++) {
      const passage = passages[p];
      const position = runPositions[passage.startRunIndex];
      if (!position) continue;
      const block = blocks[position.blockIndex];
      syncInfographicsByLine(block);
      const lineIndex = lineIndexFromCharOffset(block.text, position.charIndex);
      block.infographicsByLine[lineIndex].push(passage.text);
    }

    return blocks;
  }

  function commitBlock(chapter, block) {
    const last = chapter.blocks[chapter.blocks.length - 1];
    if (
      block.type === "camera" &&
      last &&
      last.type === "camera" &&
      last.camera === block.camera
    ) {
      last.text += "\n" + block.text;
      last.infographicsByLine = last.infographicsByLine.concat(
        cloneInfographicsByLine(block.infographicsByLine)
      );
      syncInfographicsByLine(last);
    } else {
      chapter.blocks.push(block);
    }
  }

  function appendParagraphBlocks(chapter, runs) {
    const blocks = runsToBlocks(runs);
    let pending = null;
    for (let i = 0; i < blocks.length; i++) {
      if (blocks[i].text === "") continue;
      if (pending && pending.type === "camera" && pending.camera === blocks[i].camera) {
        pending.infographicsByLine[0] = pending.infographicsByLine[0].concat(
          blocks[i].infographicsByLine[0]
        );
        for (let li = 1; li < blocks[i].infographicsByLine.length; li++) {
          pending.infographicsByLine.push(blocks[i].infographicsByLine[li].slice());
        }
        pending.text += blocks[i].text;
        syncInfographicsByLine(pending);
      } else {
        if (pending) commitBlock(chapter, pending);
        pending = {
          type: "camera",
          camera: blocks[i].camera,
          text: blocks[i].text,
          infographicsByLine: cloneInfographicsByLine(blocks[i].infographicsByLine)
        };
      }
    }
    if (pending) commitBlock(chapter, pending);
  }

  function isScriptEndMarker(text) {
    const trimmed = String(text || "").trim();
    if (trimmed.toLowerCase() === "fin du script") return true;
    const normalized = trimmed
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toUpperCase();
    return normalized.startsWith("RECAP CAM");
  }

  function buildDocumentModel(paragraphs) {
    const model = { title: "", chapters: [] };
    let index = 0;
    while (index < paragraphs.length && paragraphText(paragraphs[index]).trim() === "") index += 1;
    if (index < paragraphs.length) {
      model.title = paragraphText(paragraphs[index]).trim();
      index += 1;
    }

    let current = null;
    for (let i = index; i < paragraphs.length; i++) {
      const paragraph = paragraphs[i];
      const trimmed = paragraphText(paragraph).trim();
      if (isScriptEndMarker(trimmed)) break;
      if (isMetaParagraph(paragraph)) continue;
      if (trimmed === "") continue;
      if (paragraph.paragraphBold && paragraph.paragraphSize === 24) {
        current = { title: trimmed, blocks: [] };
        model.chapters.push(current);
        continue;
      }
      if (!current) {
        current = { title: "SANS TITRE", blocks: [] };
        model.chapters.push(current);
      }
      if (isInsertParagraph(paragraph)) {
        current.blocks.push({ type: "insert", text: trimmed });
        continue;
      }
      appendParagraphBlocks(current, paragraph.runs);
    }
    return model;
  }

  function estimateSeconds(text, wpm) {
    const words = String(text || "").replace(/\n/g, " ").split(" ").filter(function (word) {
      return word !== "";
    });
    const rate = Number(wpm);
    if (!Number.isFinite(rate) || rate <= 0) return 0;
    return Math.round((words.length / rate) * 60);
  }

  function formatDuration(seconds) {
    const total = Math.max(0, Math.round(Number(seconds) || 0));
    const minutes = Math.floor(total / 60);
    const remain = total % 60;
    return minutes + ":" + (remain < 10 ? "0" : "") + remain;
  }

  function chapterPlainText(chapter) {
    let text = "";
    for (let i = 0; i < chapter.blocks.length; i++) {
      const block = chapter.blocks[i];
      if (block.type === "insert") continue;
      if (text) text += " ";
      text += block.text;
    }
    return text;
  }

  function setDownloadActionsVisible(visible) {
    downloadActions.hidden = !visible;
    btnDownloadPrompteur.disabled = !visible;
    btnDownloadFiche.disabled = !visible;
    btnDownloadBoth.disabled = !visible;
  }

  function sanitizeTitle(title) {
    const sanitized = String(title || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toUpperCase()
      .replace(/\s+/g, "_")
      .replace(/[^A-Z0-9_]/g, "");
    return sanitized || "SANS_TITRE";
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    link.click();
    URL.revokeObjectURL(url);
  }

  function delay(ms) {
    return new Promise(function (resolve) {
      setTimeout(resolve, ms);
    });
  }

  async function downloadPrompteur(model) {
    const blob = await buildPrompteurDocx(model);
    downloadBlob(blob, "SCRIPT_PROMPTEUR_" + sanitizeTitle(model.title) + ".docx");
  }

  async function downloadFiche(model) {
    const blob = await buildFichePostprodDocx(model, wpmInput.value);
    downloadBlob(blob, "FICHE_POSTPROD_" + sanitizeTitle(model.title) + ".docx");
  }

  async function downloadBoth(model) {
    await downloadPrompteur(model);
    await delay(300);
    await downloadFiche(model);
  }

  btnDownloadPrompteur.addEventListener("click", function () {
    if (!currentModel) return;
    btnDownloadPrompteur.disabled = true;
    downloadPrompteur(currentModel)
      .catch(function (error) {
        console.error(error);
      })
      .finally(function () {
        if (currentModel) btnDownloadPrompteur.disabled = false;
      });
  });

  btnDownloadFiche.addEventListener("click", function () {
    if (!currentModel) return;
    btnDownloadFiche.disabled = true;
    downloadFiche(currentModel)
      .catch(function (error) {
        console.error(error);
      })
      .finally(function () {
        if (currentModel) btnDownloadFiche.disabled = false;
      });
  });

  btnDownloadBoth.addEventListener("click", function () {
    if (!currentModel) return;
    btnDownloadBoth.disabled = true;
    downloadBoth(currentModel)
      .catch(function (error) {
        console.error(error);
      })
      .finally(function () {
        if (currentModel) btnDownloadBoth.disabled = false;
      });
  });

  function showMessage(message) {
    currentModel = null;
    setDownloadActionsVisible(false);
    resultsOutput.replaceChildren();
    const p = document.createElement("p");
    p.className = "results-message";
    p.textContent = message;
    resultsOutput.appendChild(p);
  }

  function renderModel(model) {
    const wpm = wpmInput.value;
    let totalSeconds = 0;
    const durations = [];
    for (let i = 0; i < model.chapters.length; i++) {
      const seconds = estimateSeconds(chapterPlainText(model.chapters[i]), wpm);
      durations.push(seconds);
      totalSeconds += seconds;
    }

    resultsOutput.replaceChildren();
    setDownloadActionsVisible(true);
    const total = document.createElement("p");
    total.textContent = "Durée totale estimée : " + formatDuration(totalSeconds);
    resultsOutput.appendChild(total);
    if (model.title) {
      const title = document.createElement("h2");
      title.textContent = model.title.toUpperCase();
      resultsOutput.appendChild(title);
    }
    for (let i = 0; i < model.chapters.length; i++) {
      const chapter = model.chapters[i];
      const heading = document.createElement("h2");
      heading.textContent = chapter.title.toUpperCase() + " (" + formatDuration(durations[i]) + ")";
      resultsOutput.appendChild(heading);
      const list = document.createElement("ul");
      for (let j = 0; j < chapter.blocks.length; j++) {
        const block = chapter.blocks[j];
        const li = document.createElement("li");
        const label = document.createElement("strong");
        label.textContent = block.type === "insert" ? "INSERT VIDÉO" : "CAMERA " + block.camera;
        li.appendChild(label);
        li.appendChild(document.createElement("br"));
        li.appendChild(document.createTextNode(block.text));
        list.appendChild(li);
      }
      resultsOutput.appendChild(list);
    }
  }

  async function readDocx(file) {
    if (!file) {
      showMessage("Aucun fichier reçu.");
      return;
    }
    if (!file.name.toLowerCase().endsWith(".docx")) {
      showMessage("Ce fichier n'est pas un .docx.");
      return;
    }

    try {
      const buffer = await file.arrayBuffer();
      const zip = await JSZip.loadAsync(buffer);
      const entry = zip.file("word/document.xml");
      if (!entry) {
        showMessage("Document Word introuvable dans ce fichier.");
        return;
      }
      const xml = await entry.async("string");
      const paragraphs = parseDocumentXml(xml);
      console.log(paragraphs);
      const model = buildDocumentModel(paragraphs);
      console.log(model);
      currentModel = model;
      renderModel(model);
    } catch (error) {
      console.error(error);
      showMessage("Lecture du script impossible.");
    }
  }
})();
