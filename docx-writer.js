(function (global) {
  "use strict";

  var W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  var RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
  var CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
  var OFFICE_DOC_REL =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
  var STYLES_REL =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
  var SETTINGS_REL =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";

  function escapeXml(text) {
    return String(text)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  function textTag(text) {
    var value = String(text);
    var preserve = /^\s|\s$/.test(value) ? ' xml:space="preserve"' : "";
    return "<w:t" + preserve + ">" + escapeXml(value) + "</w:t>";
  }

  function run(text, rPr) {
    return "<w:r>" + (rPr || "") + textTag(text) + "</w:r>";
  }

  function rPr(options) {
    var parts = ["<w:rPr>"];
    if (options.bold) parts.push("<w:b/>");
    if (options.italic) parts.push("<w:i/>");
    if (options.underline) parts.push('<w:u w:val="single"/>');
    if (options.size != null) parts.push('<w:sz w:val="' + options.size + '"/>');
    parts.push("</w:rPr>");
    return parts.join("");
  }

  function spacingAttrs(spacing) {
    if (!spacing) return "";
    var attrs = [];
    if (spacing.before != null) attrs.push('w:before="' + spacing.before + '"');
    if (spacing.after != null) attrs.push('w:after="' + spacing.after + '"');
    if (spacing.line != null) attrs.push('w:line="' + spacing.line + '"');
    if (spacing.lineRule != null) attrs.push('w:lineRule="' + spacing.lineRule + '"');
    return attrs.length ? " " + attrs.join(" ") : "";
  }

  function paragraph(content, options) {
    options = options || {};
    var pPr = "";
    if (options.spacing) {
      pPr = "<w:pPr><w:spacing" + spacingAttrs(options.spacing) + "/></w:pPr>";
    }
    return "<w:p>" + pPr + (content || "") + "</w:p>";
  }

  function emptyParagraph(spacingAfter) {
    if (spacingAfter == null) return "<w:p/>";
    return paragraph("", { spacing: { after: spacingAfter } });
  }

  function styledParagraph(text, style, spacing) {
    return paragraph(run(text, rPr(style)), { spacing: spacing });
  }

  function splitParagraphs(text) {
    return String(text || "").split(/\r?\n/);
  }

  function sectionProperties() {
    return (
      "<w:sectPr>" +
      '<w:pgSz w:w="12240" w:h="15840"/>' +
      '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" ' +
      'w:header="720" w:footer="720" w:gutter="0"/>' +
      "</w:sectPr>"
    );
  }

  function wrapDocument(bodyContent) {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="' +
      W_NS +
      '">' +
      "<w:body>" +
      bodyContent +
      sectionProperties() +
      "</w:body>" +
      "</w:document>"
    );
  }

  function tableCell(text, options) {
    options = options || {};
    var cellText = text == null ? "" : String(text);
    var style = options.style || { size: 18 };
    var tcPr =
      '<w:tcPr><w:tcW w:w="2880" w:type="dxa"/></w:tcPr>';
    return (
      "<w:tc>" +
      tcPr +
      styledParagraph(cellText, style, null) +
      "</w:tc>"
    );
  }

  function tableRow(cells) {
    return "<w:tr>" + cells.join("") + "</w:tr>";
  }

  function postprodTable(rows) {
    var grid =
      "<w:tblGrid>" +
      '<w:gridCol w:w="2880"/>' +
      '<w:gridCol w:w="2880"/>' +
      '<w:gridCol w:w="2880"/>' +
      "</w:tblGrid>";
    var tblPr =
      "<w:tblPr>" +
      '<w:tblW w:w="8640" w:type="dxa"/>' +
      '<w:tblLayout w:type="fixed"/>' +
      "<w:tblBorders>" +
      '<w:top w:val="single" w:sz="4" w:color="auto"/>' +
      '<w:left w:val="single" w:sz="4" w:color="auto"/>' +
      '<w:bottom w:val="single" w:sz="4" w:color="auto"/>' +
      '<w:right w:val="single" w:sz="4" w:color="auto"/>' +
      '<w:insideH w:val="single" w:sz="4" w:color="auto"/>' +
      '<w:insideV w:val="single" w:sz="4" w:color="auto"/>' +
      "</w:tblBorders>" +
      "</w:tblPr>";
    return "<w:tbl>" + tblPr + grid + rows.join("") + "</w:tbl>";
  }

  function chapterParagraphs(chapter) {
    var parts = [];
    for (var i = 0; i < chapter.blocks.length; i++) {
      var block = chapter.blocks[i];
      var lines = splitParagraphs(block.text);
      for (var j = 0; j < lines.length; j++) {
        if (block.type === "insert") {
          parts.push("[INSERT VIDÉO] " + lines[j]);
        } else {
          parts.push(lines[j]);
        }
      }
    }
    return parts;
  }

  function buildPrompteurBody(documentModel) {
    var parts = [];

    parts.push(
      styledParagraph("SCRIPT TOURNAGE — PROMPTEUR", { bold: true, size: 44 }, { after: 120 })
    );
    parts.push(
      styledParagraph(documentModel.title || "", { bold: true, size: 32 }, { after: 120 })
    );
    parts.push(
      styledParagraph("Storytelling / Décryptage", { italic: true, size: 22 }, null)
    );
    parts.push(emptyParagraph(80));

    for (var c = 0; c < documentModel.chapters.length; c++) {
      var chapter = documentModel.chapters[c];
      parts.push(
        styledParagraph(
          chapter.title || "",
          { bold: true, underline: true, size: 32 },
          { before: 480, after: 120 }
        )
      );

      for (var b = 0; b < chapter.blocks.length; b++) {
        var block = chapter.blocks[b];
        if (block.type === "insert") {
          parts.push(
            styledParagraph(
              "INSERT VIDÉO",
              { bold: true, size: 26 },
              { before: 200, after: 80 }
            )
          );
          var insertLines = splitParagraphs(block.text);
          for (var il = 0; il < insertLines.length; il++) {
            parts.push(
              styledParagraph(
                insertLines[il],
                { italic: true, size: 28 },
                { after: 160, line: 312, lineRule: "auto" }
              )
            );
          }
        } else {
          var cameraLabel = "CAMERA " + block.camera;
          parts.push(
            styledParagraph(
              cameraLabel,
              { bold: true, size: 26 },
              { before: 200, after: 80 }
            )
          );

          var lines = splitParagraphs(block.text);
          for (var l = 0; l < lines.length; l++) {
            parts.push(
              styledParagraph(lines[l], { size: 28 }, { after: 160, line: 312, lineRule: "auto" })
            );
          }
        }
      }
    }

    return parts.join("");
  }

  function buildFichePostprodBody(documentModel) {
    var parts = [];

    parts.push(
      styledParagraph(
        "FICHE POST-PRODUCTION — SCRIPT COMPLET",
        { bold: true, size: 40 },
        null
      )
    );
    parts.push(
      styledParagraph(documentModel.title || "", { bold: true, size: 28 }, null)
    );
    parts.push(
      styledParagraph(
        "Une ligne par phrase/paragraphe du script. Cases laissées vides quand l'information n'est pas explicitement présente dans le texte.",
        { italic: true, size: 18 },
        null
      )
    );
    parts.push(emptyParagraph(null));

    for (var c = 0; c < documentModel.chapters.length; c++) {
      var chapter = documentModel.chapters[c];
      parts.push(
        styledParagraph(chapter.title || "", { bold: true, size: 26 }, null)
      );

      var rows = [];
      rows.push(
        tableRow([
          tableCell("Phrase repère (extrait du script)", { style: { bold: true, size: 18 } }),
          tableCell("Incrustation / image à prévoir", { style: { bold: true, size: 18 } }),
          tableCell("Sources", { style: { bold: true, size: 18 } })
        ])
      );

      var paragraphs = chapterParagraphs(chapter);
      for (var p = 0; p < paragraphs.length; p++) {
        rows.push(
          tableRow([
            tableCell(paragraphs[p], { style: { size: 18 } }),
            tableCell("", { style: { size: 18 } }),
            tableCell("", { style: { size: 18 } })
          ])
        );
      }

      parts.push(postprodTable(rows));
      parts.push(emptyParagraph(null));
    }

    return parts.join("");
  }

  function contentTypesXml() {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="' +
      CT_NS +
      '">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
      "</Types>"
    );
  }

  function packageRelsXml() {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="' +
      RELS_NS +
      '">' +
      '<Relationship Id="rId1" Type="' +
      OFFICE_DOC_REL +
      '" Target="word/document.xml"/>' +
      "</Relationships>"
    );
  }

  function documentRelsXml() {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="' +
      RELS_NS +
      '">' +
      '<Relationship Id="rId1" Type="' +
      STYLES_REL +
      '" Target="styles.xml"/>' +
      '<Relationship Id="rId2" Type="' +
      SETTINGS_REL +
      '" Target="settings.xml"/>' +
      "</Relationships>"
    );
  }

  function stylesXml() {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:w="' +
      W_NS +
      '">' +
      "<w:docDefaults>" +
      "<w:rPrDefault><w:rPr/></w:rPrDefault>" +
      "<w:pPrDefault/>" +
      "</w:docDefaults>" +
      "</w:styles>"
    );
  }

  function settingsXml() {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:settings xmlns:w="' +
      W_NS +
      '"></w:settings>'
    );
  }

  function assembleDocx(documentXml) {
    if (typeof JSZip === "undefined") {
      return Promise.reject(new Error("JSZip indisponible"));
    }

    var zip = new JSZip();
    zip.file(
      "mimetype",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      { compression: "STORE" }
    );
    zip.file("[Content_Types].xml", contentTypesXml());
    zip.folder("_rels").file(".rels", packageRelsXml());
    zip.folder("word").file("document.xml", documentXml);
    zip.folder("word").folder("_rels").file("document.xml.rels", documentRelsXml());
    zip.folder("word").file("styles.xml", stylesXml());
    zip.folder("word").file("settings.xml", settingsXml());

    return zip.generateAsync({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      compression: "DEFLATE",
      compressionOptions: { level: 6 }
    });
  }

  function buildPrompteurDocx(documentModel) {
    var body = buildPrompteurBody(documentModel || { title: "", chapters: [] });
    return assembleDocx(wrapDocument(body));
  }

  function buildFichePostprodDocx(documentModel) {
    var body = buildFichePostprodBody(documentModel || { title: "", chapters: [] });
    return assembleDocx(wrapDocument(body));
  }

  global.buildPrompteurDocx = buildPrompteurDocx;
  global.buildFichePostprodDocx = buildFichePostprodDocx;
})(typeof window !== "undefined" ? window : globalThis);
