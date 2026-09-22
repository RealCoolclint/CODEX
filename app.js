(function () {
  const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

  const dropzone = document.getElementById("dropzone");
  const results = document.getElementById("results");

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
        highlight: rPr ? readHighlight(rPr) : null
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

  function formatPoints(halfPoints) {
    const points = halfPoints / 2;
    const label = Number.isInteger(points) ? String(points) : String(points);
    return label + "pt";
  }

  function formatAttrs(paragraph) {
    const parts = [];
    if (paragraph.paragraphBold) parts.push("gras");
    if (typeof paragraph.paragraphSize === "number") parts.push(formatPoints(paragraph.paragraphSize));
    return parts.length ? "[" + parts.join(", ") + "] " : "";
  }

  function showMessage(message) {
    results.replaceChildren();
    const p = document.createElement("p");
    p.className = "results-message";
    p.textContent = message;
    results.appendChild(p);
  }

  function renderParagraphs(paragraphs) {
    results.replaceChildren();
    const list = document.createElement("ul");
    for (let i = 0; i < paragraphs.length; i++) {
      const paragraph = paragraphs[i];
      const li = document.createElement("li");
      const attrs = formatAttrs(paragraph);
      if (attrs) {
        const span = document.createElement("span");
        span.className = "para-attrs";
        span.textContent = attrs;
        li.appendChild(span);
      }
      for (let j = 0; j < paragraph.runs.length; j++) {
        const run = paragraph.runs[j];
        if (run.highlight) {
          const mark = document.createElement("span");
          mark.className = "para-attrs";
          mark.textContent = "[SURLIGNÉ] ";
          li.appendChild(mark);
        }
        li.appendChild(document.createTextNode(run.text));
      }
      list.appendChild(li);
    }
    results.appendChild(list);
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
      renderParagraphs(paragraphs);
    } catch (error) {
      console.error(error);
      showMessage("Lecture du script impossible.");
    }
  }
})();
