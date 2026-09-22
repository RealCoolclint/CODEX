(function () {
  const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

  const dropzone = document.getElementById("dropzone");
  const results = document.getElementById("results");
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

  function parseParagraph(paragraph) {
    const runNodes = paragraph.getElementsByTagNameNS(W_NS, "r");
    const runs = [];
    let paragraphBold = false;
    let paragraphSize = null;
    let representativeFound = false;

    for (let i = 0; i < runNodes.length; i++) {
      const run = runNodes[i];
      if (!belongsTo(run, paragraph)) continue;

      const rPr = directChild(run, "rPr");
      let text = "";
      const texts = run.getElementsByTagNameNS(W_NS, "t");
      for (let j = 0; j < texts.length; j++) {
        if (!belongsTo(texts[j], paragraph)) continue;
        text += texts[j].textContent || "";
      }

      runs.push({
        text: text,
        highlight: rPr ? readHighlight(rPr) : null,
        italic: rPr ? isItalic(rPr) : false,
        color: rPr ? readColor(rPr) : null
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

  function runsToBlocks(runs) {
    const blocks = [];
    let leading = "";
    for (let i = 0; i < runs.length; i++) {
      const text = runs[i].text;
      if (text.trim() === "") {
        if (blocks.length) blocks[blocks.length - 1].text += text;
        else leading += text;
        continue;
      }
      const camera = runs[i].highlight ? 2 : 1;
      const last = blocks[blocks.length - 1];
      if (last && last.camera === camera) last.text += text;
      else {
        blocks.push({ camera: camera, text: leading + text });
        leading = "";
      }
    }
    return blocks;
  }

  function commitBlock(chapter, block) {
    const last = chapter.blocks[chapter.blocks.length - 1];
    if (last && last.camera === block.camera) last.text += "\n" + block.text;
    else chapter.blocks.push(block);
  }

  function appendParagraphBlocks(chapter, runs) {
    const blocks = runsToBlocks(runs);
    let pending = null;
    for (let i = 0; i < blocks.length; i++) {
      if (blocks[i].text === "") continue;
      if (pending && pending.camera === blocks[i].camera) pending.text += blocks[i].text;
      else {
        if (pending) commitBlock(chapter, pending);
        pending = { camera: blocks[i].camera, text: blocks[i].text };
      }
    }
    if (pending) commitBlock(chapter, pending);
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
      if (trimmed.toLowerCase() === "fin du script") break;
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
      if (text) text += " ";
      text += chapter.blocks[i].text;
    }
    return text;
  }

  function showMessage(message) {
    currentModel = null;
    results.replaceChildren();
    const p = document.createElement("p");
    p.className = "results-message";
    p.textContent = message;
    results.appendChild(p);
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

    results.replaceChildren();
    const total = document.createElement("p");
    total.textContent = "Durée totale estimée : " + formatDuration(totalSeconds);
    results.appendChild(total);
    if (model.title) {
      const title = document.createElement("h2");
      title.textContent = model.title.toUpperCase();
      results.appendChild(title);
    }
    for (let i = 0; i < model.chapters.length; i++) {
      const chapter = model.chapters[i];
      const heading = document.createElement("h2");
      heading.textContent = chapter.title.toUpperCase() + " (" + formatDuration(durations[i]) + ")";
      results.appendChild(heading);
      const list = document.createElement("ul");
      for (let j = 0; j < chapter.blocks.length; j++) {
        const block = chapter.blocks[j];
        const li = document.createElement("li");
        const label = document.createElement("strong");
        label.textContent = "CAMERA " + block.camera;
        li.appendChild(label);
        li.appendChild(document.createElement("br"));
        li.appendChild(document.createTextNode(block.text));
        list.appendChild(li);
      }
      results.appendChild(list);
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
