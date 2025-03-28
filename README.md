(() => {
  "use strict";

  // Konfiguracja i globalne zmienne
  const cfg = {
    storageKey: "aggregatedResults",
    excelFlagKey: "triggerExcelAfterReload",
    bannedNIPs: ["0000116894", "0000681163", "7010046065"],
    bannedTels: [
      "012110943", "140756502", "120904481", "150576738",
      "150576742", "150576746", "150576750", "150576754",
      "150576758", "150573522", "150576732"
    ],
    bannedEmails: ["iod@wenet.pl", "kontakt@wenet.pl"],
    reNIP: /\b\d{3}[- ]?\d{3}[- ]?\d{2}[- ]?\d{2}\b/g,
    reTel: /(?:\(\d{2}\)\s*\d{3}[-\s]?\d{2}[-\s]?\d{2})|\b\d{3}[- ]?\d{3}[- ]?\d{3}\b/g,
    reEmail: /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-z]{2,}/gi,
    refresh: 1000,
    defaultEnabledColumns: {
      url: true,
      companyName: true,
      NIPs: true,
      telefony: true,
      emaile: true,
      timestamp: true
    },
    columnsKey: "enabledColumns"
  };

  let jsonData = localStorage.getItem("excelData")
    ? JSON.parse(localStorage.getItem("excelData"))
    : null;
  let currentRowIndex = 1;
  let autoSearchTriggered = false;
  let lastSearchTime = Date.now();

  // Funkcje pomocnicze – selektory i style
  const qs = (sel, parent = document) => parent.querySelector(sel);
  const qsa = (sel, parent = document) => Array.from(parent.querySelectorAll(sel));
  const injectStyle = (id, css) => {
    if (!document.getElementById(id)) {
      const styleEl = document.createElement("style");
      styleEl.id = id;
      styleEl.textContent = css;
      document.head.append(styleEl);
    }
  };
  const setStyles = (el, styles) => Object.assign(el.style, styles);
  const shortenText = (text, maxLength = 50) =>
    typeof text === "string" && text.length > maxLength ? text.slice(0, maxLength) + "..." : text;

  // Funkcja umożliwiająca przeciąganie elementów (z pominięciem przycisków)
  const makeDraggable = (el) => {
    let pos = { x: 0, y: 0, startX: 0, startY: 0 };
    el.onmousedown = (e) => {
      if (e.target.tagName === "BUTTON") return;
      e.preventDefault();
      pos.startX = e.clientX;
      pos.startY = e.clientY;
      document.onmousemove = (evt) => {
        evt.preventDefault();
        pos.x = pos.startX - evt.clientX;
        pos.y = pos.startY - evt.clientY;
        pos.startX = evt.clientX;
        pos.startY = evt.clientY;
        el.style.top = el.offsetTop - pos.y + "px";
        el.style.left = el.offsetLeft - pos.x + "px";
      };
      document.onmouseup = () => {
        document.onmouseup = document.onmousemove = null;
      };
    };
  };

  // Obsługa localStorage – dane i konfiguracja kolumn
  const getData = () => {
    try {
      return JSON.parse(localStorage.getItem(cfg.storageKey)) || [];
    } catch (e) {
      console.error("Error reading data", e);
      return [];
    }
  };
  const saveData = (data) => localStorage.setItem(cfg.storageKey, JSON.stringify(data));
  const getEnabledColumns = () => {
    try {
      const stored = localStorage.getItem(cfg.columnsKey);
      return stored ? JSON.parse(stored) : { ...cfg.defaultEnabledColumns };
    } catch (e) {
      console.error("Error reading columns config", e);
      return { ...cfg.defaultEnabledColumns };
    }
  };
  const saveEnabledColumns = (cols) => localStorage.setItem(cfg.columnsKey, JSON.stringify(cols));

  // Funkcja zbierająca dane z dokumentu
  const collectData = () => {
    const bodyClone = document.body.cloneNode(true);
    const mapElem = bodyClone.querySelector("#map");
    if (mapElem) mapElem.remove();
    const bodyText = bodyClone.textContent;
    const pageTitle = qs("h1")?.textContent.trim() || document.title.trim();

    if (pageTitle.toLowerCase() === "wyniki dla fraz") {
      console.log("Wyniki dla fraz - pomijam.");
      return null;
    }

    const formatMatches = (regex, banned = [], cleanFn = (str) => str.replace(/[-\s()]/g, "")) =>
      [...new Set((bodyText.match(regex) || []).filter(item => !banned.includes(cleanFn(item))))];

    const nips = formatMatches(cfg.reNIP, cfg.bannedNIPs);
    const tels = formatMatches(cfg.reTel, cfg.bannedTels, (t) => t.replace(/[()\-\s]/g, ""))
                  .filter(t => t.charAt(0) !== "0").slice(0, 1);
    const emails = formatMatches(cfg.reEmail, cfg.bannedEmails);

    return {
      url: location.href,
      companyName: pageTitle,
      NIPs: nips,
      telefony: tels,
      emaile: emails,
      timestamp: new Date().toISOString()
    };
  };

  const isDuplicate = (newRecord, records) => {
    return records.some(record => {
      if (record.url === newRecord.url) return true;
      if (newRecord.NIPs.length && record.NIPs.length) {
        return newRecord.NIPs.some(nip => record.NIPs.includes(nip));
      }
      return false;
    });
  };

  // Funkcja generująca plik Excel
  const generateExcel = () => {
    const doGen = () => {
      try {
        const data = getData();
        const enabled = getEnabledColumns();
        const header = Object.keys(enabled)
          .filter(col => enabled[col])
          .map(col => {
            switch (col) {
              case "url": return "Page URL";
              case "companyName": return "Nazwa firmy";
              case "NIPs": return "NIPs";
              case "telefony": return "Telefony";
              case "emaile": return "E-maile";
              case "timestamp": return "Timestamp";
              default: return col;
            }
          });
        const rows = data.map(item =>
          Object.keys(enabled).filter(col => enabled[col])
            .map(col => Array.isArray(item[col]) ? item[col].join(", ") : item[col])
        );
        const sheetData = [header, ...rows];
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sheetData), "Aggregated Data");
        XLSX.writeFile(wb, "pobrane_wyniki.xlsx");
        console.log("Excel wygenerowany.");
        localStorage.setItem(cfg.excelFlagKey, "true");
        saveData([]);
        updateUIPanel();
      } catch (e) {
        console.error("Excel generation error", e);
      }
    };

    if (typeof XLSX === "undefined") {
      const s = document.createElement("script");
      s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
      s.onload = doGen;
      s.onerror = () => console.error("Error loading XLSX library");
      document.head.append(s);
    } else {
      doGen();
    }
  };

  // Funkcje UI – modal, tabele, przyciski
  const createModal = (id, titleText) => {
    let modal = qs(`#${id}`);
    if (!modal) {
      modal = document.createElement("div");
      modal.id = id;
      setStyles(modal, {
        position: "fixed", top: "0", left: "0", width: "100%",
        height: "100%", backgroundColor: "rgba(0,0,0,0.5)",
        zIndex: "10000", display: "flex", justifyContent: "center", alignItems: "center"
      });
      const content = document.createElement("div");
      content.id = `${id}Content`;
      setStyles(content, {
        backgroundColor: "#fff", padding: "20px", borderRadius: "8px",
        maxHeight: "80%", overflowY: "auto", width: "90%", maxWidth: "800px"
      });
      modal.appendChild(content);
      document.body.append(modal);
      // Przycisk zamknięcia
      const closeBtn = document.createElement("button");
      closeBtn.textContent = "Zamknij";
      setStyles(closeBtn, {
        backgroundColor: "#dc3545", color: "#fff", border: "none",
        borderRadius: "4px", padding: "8px 12px", cursor: "pointer", marginBottom: "10px"
      });
      closeBtn.addEventListener("click", () => { modal.style.display = "none"; });
      content.appendChild(closeBtn);
      // Nagłówek, jeśli przekazano tytuł
      if (titleText) {
        const header = document.createElement("h3");
        header.textContent = titleText;
        content.appendChild(header);
      }
    } else {
      modal.style.display = "flex";
    }
    return modal;
  };

  const showPreview = () => {
    const modal = createModal("aggregatedPreviewModal", "Podgląd wyników");
    const content = qs("#aggregatedPreviewModalContent", modal) || qs("#aggregatedPreviewModalContent", modal.firstElementChild);
    // Usuwamy wszystkie dodatkowe elementy, zachowując przycisk zamknięcia
    while (content.childNodes.length > 1) content.removeChild(content.lastChild);
    
    const data = getData();
    const enabled = getEnabledColumns();
    if (!data.length) {
      const p = document.createElement("p");
      p.textContent = "Brak wyników do wyświetlenia.";
      content.appendChild(p);
      return;
    }
    const table = document.createElement("table");
    setStyles(table, { width: "100%", borderCollapse: "collapse", marginTop: "10px" });
    // Nagłówek tabeli
    const headerRow = document.createElement("tr");
    const emptyTh = document.createElement("th");
    emptyTh.textContent = "";
    headerRow.appendChild(emptyTh);
    Object.keys(enabled).forEach(col => {
      if (enabled[col]) {
        const th = document.createElement("th");
        th.textContent = (col === "url") ? "Page URL" :
                         (col === "companyName") ? "Nazwa firmy" :
                         (col === "NIPs") ? "NIPs" :
                         (col === "telefony") ? "Telefony" :
                         (col === "emaile") ? "E-maile" :
                         (col === "timestamp") ? "Timestamp" : col;
        setStyles(th, { border: "1px solid #ddd", padding: "8px", backgroundColor: "#f2f2f2" });
        headerRow.appendChild(th);
      }
    });
    table.appendChild(headerRow);
    // Wiersze danych
    data.forEach((item, index) => {
      const row = document.createElement("tr");
      // Przycisk usuwania
      const deleteTd = document.createElement("td");
      const deleteBtn = document.createElement("button");
      deleteBtn.textContent = "x";
      setStyles(deleteBtn, {
        cursor: "pointer", backgroundColor: "#dc3545", color: "#fff",
        border: "none", borderRadius: "4px", padding: "2px 6px"
      });
      deleteBtn.addEventListener("click", () => {
        const allData = getData();
        allData.splice(index, 1);
        saveData(allData);
        updateUIPanel();
        showPreview();
      });
      deleteTd.appendChild(deleteBtn);
      setStyles(deleteTd, { border: "1px solid #ddd", padding: "8px" });
      row.appendChild(deleteTd);
      // Dane kolumn
      Object.keys(enabled).forEach(col => {
        if (enabled[col]) {
          const td = document.createElement("td");
          let value = item[col];
          if (col === "url" && typeof value === "string") {
            td.title = value;
            value = shortenText(value);
          }
          td.textContent = Array.isArray(value) ? value.join(", ") : value;
          setStyles(td, { border: "1px solid #ddd", padding: "8px" });
          row.appendChild(td);
        }
      });
      table.appendChild(row);
    });
    content.appendChild(table);
  };

  const showExcelPreview = () => {
    const modal = createModal("excelPreviewModal", "Podgląd Excela");
    const content = qs("#excelPreviewModalContent", modal) || qs("#excelPreviewModalContent", modal.firstElementChild);
    // Usuwamy dodatkową zawartość poza przyciskiem zamknięcia
    while (content.childNodes.length > 1) content.removeChild(content.lastChild);
    // Sekcja wyboru pliku
    if (!qs("#excelFileSelector", content)) {
      const fileSelectorDiv = document.createElement("div");
      fileSelectorDiv.id = "excelFileSelector";
      fileSelectorDiv.style.marginBottom = "15px";
      const fileLabel = document.createElement("label");
      fileLabel.textContent = "Wybierz plik Excel: ";
      const fileInput = document.createElement("input");
      fileInput.type = "file";
      fileInput.accept = ".xlsx, .xls";
      fileInput.addEventListener("change", (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (evt) => {
          const data = evt.target.result;
          const loadExcelData = (data) => {
            const workbook = XLSX.read(data, { type: "binary" });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            // Zapisujemy dane Excel do localStorage, aby autoUpdateSearchAndClick mogło je pobrać
            localStorage.setItem("excelData", JSON.stringify(jsonData));
            renderExcelTable();
          };
          if (typeof XLSX === "undefined") {
            const s = document.createElement("script");
            s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
            s.onload = () => loadExcelData(data);
            s.onerror = () => console.error("Error loading XLSX library");
            document.head.append(s);
          } else {
            loadExcelData(data);
          }
        };
        reader.readAsBinaryString(file);
      });
      fileLabel.appendChild(fileInput);
      fileSelectorDiv.appendChild(fileLabel);
      content.appendChild(fileSelectorDiv);
    }
    // Renderowanie tabeli Excel
    const renderExcelTable = () => {
      let tableContainer = qs("#excelPreviewTableContainer", content);
      if (!tableContainer) {
        tableContainer = document.createElement("div");
        tableContainer.id = "excelPreviewTableContainer";
        content.appendChild(tableContainer);
      }
      tableContainer.innerHTML = "";
      if (!jsonData || jsonData.length === 0) {
        tableContainer.textContent = "Brak danych Excela.";
        return;
      }
      const table = document.createElement("table");
      setStyles(table, { width: "100%", borderCollapse: "collapse", marginTop: "10px" });
      // Nagłówek tabeli
      const headerRow = document.createElement("tr");
      jsonData[0].forEach(header => {
        const th = document.createElement("th");
        th.textContent = header === "Company Name" ? "Nazwa firmy" : header;
        setStyles(th, { border: "1px solid #ddd", padding: "8px", backgroundColor: "#f2f2f2" });
        headerRow.appendChild(th);
      });
      table.appendChild(headerRow);
      // Wiersze danych
      for (let i = 1; i < jsonData.length; i++) {
        const row = document.createElement("tr");
        jsonData[i].forEach(cell => {
          const td = document.createElement("td");
          td.textContent = cell;
          setStyles(td, { border: "1px solid #ddd", padding: "8px" });
          row.appendChild(td);
        });
        table.appendChild(row);
      }
      tableContainer.appendChild(table);
    };
    renderExcelTable();
  };

  const openColumnsModal = () => {
    const modal = createModal("columnsSelectionModal", "Wybierz kolumny do zbierania");
    const content = qs("#columnsSelectionModalContent", modal) || modal.firstElementChild;
    // Usuwamy dodatkowe elementy, zachowując przycisk zamknięcia
    while (content.childNodes.length > 1) content.removeChild(content.lastChild);
    const header = document.createElement("h3");
    header.textContent = "Wybierz kolumny do zbierania";
    content.appendChild(header);
    const form = document.createElement("form");
    const enabled = getEnabledColumns();
    Object.keys(cfg.defaultEnabledColumns).forEach(col => {
      const label = document.createElement("label");
      label.style.display = "block";
      label.style.marginBottom = "8px";
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.checked = enabled[col];
      checkbox.style.marginRight = "8px";
      checkbox.addEventListener("change", () => {
        enabled[col] = checkbox.checked;
        saveEnabledColumns(enabled);
      });
      label.appendChild(checkbox);
      label.append(col === "url" ? "Page URL" :
                   col === "companyName" ? "Nazwa firmy" :
                   col === "NIPs" ? "NIPs" :
                   col === "telefony" ? "Telefony" :
                   col === "emaile" ? "E-maile" :
                   col === "timestamp" ? "Timestamp" : col);
      form.appendChild(label);
    });
    content.appendChild(form);
  };

  // Główny panel sterowania
  const updateUIPanel = () => {
    const panel = qs("#myAggregatedResultsPanel");
    if (panel) qs("#resultCount", panel).textContent = "Ilość wyników: " + getData().length;
  };

  const createUIPanel = () => {
    if (qs("#myAggregatedResultsPanel")) return;
    const panel = document.createElement("div");
    panel.id = "myAggregatedResultsPanel";
    setStyles(panel, {
      position: "fixed", bottom: "20px", left: "20px",
      backgroundColor: "#f9f9f9", border: "1px solid #ddd",
      borderRadius: "8px", padding: "8px", zIndex: "9999",
      fontFamily: "'Segoe UI', sans-serif", fontSize: "12px",
      boxShadow: "0 4px 8px rgba(0,0,0,0.1)", maxWidth: "280px",
      maxHeight: "250px", overflowY: "auto",
      display: localStorage.getItem("panelVisible") === "true" ? "block" : "none"
    });
    const resultCount = document.createElement("span");
    resultCount.id = "resultCount";
    resultCount.textContent = "Ilość wyników: " + getData().length;
    resultCount.style.marginBottom = "8px";
    resultCount.style.display = "block";
    panel.appendChild(resultCount);

    const progressText = document.createElement("span");
    progressText.id = "progressText";
    progressText.style.display = "block";
    progressText.style.marginBottom = "8px";
    panel.appendChild(progressText);

    const createBtn = (text, bg, hoverBg, onClick, id = "") => {
      const btn = document.createElement("button");
      btn.textContent = text;
      setStyles(btn, {
        backgroundColor: bg, border: "none", borderRadius: "4px", color: "#fff",
        cursor: "pointer", padding: "8px 10px", fontSize: "12px", width: "100%",
        transition: "background-color 0.3s", marginBottom: "5px"
      });
      if (id) btn.id = id;
      btn.addEventListener("mouseover", () => btn.style.backgroundColor = hoverBg);
      btn.addEventListener("mouseout", () => btn.style.backgroundColor = bg);
      btn.addEventListener("click", onClick);
      return btn;
    };

    panel.appendChild(createBtn("Podgląd wyników", "#17a2b8", "#138496", showPreview));
    panel.appendChild(createBtn("Pobierz Excel", "#28a745", "#218838", generateExcel));
    panel.appendChild(createBtn("Podgląd Excela", "#6c757d", "#5a6268", showExcelPreview));
    panel.appendChild(createBtn("Sprawdzanie bazy danych", "#ffc107", "#ffc107", () => {
      const status = document.createElement("div");
      setStyles(status, {
        position: "fixed", top: "50%", left: "50%",
        transform: "translate(-50%, -50%)", backgroundColor: "#28a745",
        color: "#fff", padding: "20px", borderRadius: "8px", zIndex: "10000"
      });
      status.textContent = "wszystko dobrze";
      document.body.appendChild(status);
      setTimeout(() => { status.remove(); location.reload(); }, 3000);
    }));
    panel.appendChild(createBtn("Usuń wyniki", "#dc3545", "#c82333", () => {
      localStorage.removeItem(cfg.storageKey);
      updateUIPanel();
      console.log("Wszystkie wyniki usunięte.");
    }));
    panel.appendChild(createBtn("Wybierz kolumny", "#6c757d", "#5a6268", openColumnsModal));
    panel.appendChild(createBtn("Aktualizuj wynik", "#343a40", "#23272b", autoUpdateSearchAndClick, "updateResultBtn"));
    document.body.appendChild(panel);
    makeDraggable(panel);
  };

  // Przycisk do przełączania widoczności panelu
  const createToggleButton = () => {
    if (qs("#togglePanelButton")) return;
    const btn = document.createElement("button");
    btn.id = "togglePanelButton";
    btn.textContent = localStorage.getItem("panelVisible") === "true" ? "Ukryj panel" : "Pokaż panel";
    setStyles(btn, {
      position: "fixed", top: "20px", left: "20px", backgroundColor: "#007bff",
      color: "#fff", border: "none", borderRadius: "4px", padding: "10px 15px",
      cursor: "pointer", zIndex: "10000"
    });
    btn.addEventListener("click", () => {
      const panel = qs("#myAggregatedResultsPanel");
      if (panel) {
        if (panel.style.display === "none" || panel.style.display === "") {
          panel.style.display = "block";
          btn.textContent = "Ukryj panel";
          localStorage.setItem("panelVisible", "true");
        } else {
          panel.style.display = "none";
          btn.textContent = "Pokaż panel";
          localStorage.setItem("panelVisible", "false");
        }
      }
    });
    document.body.appendChild(btn);
  };

  // Automatyzacja – UI
  const setupAutomationUI = () => {
    const css = `
      .automation-container {
        position: fixed;
        top: 120px;
        left: 10px;
        z-index: 9999;
        background: rgba(255,255,255,0.95);
        border: 1px solid #ddd;
        padding: 12px;
        border-radius: 10px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        font-family: 'Segoe UI', sans-serif;
        width: 280px;
      }
      .automation-button {
        background: linear-gradient(145deg, #00aaff, #005fbb);
        color: #fff;
        border: none;
        padding: 12px 20px;
        margin: 5px 0;
        border-radius: 8px;
        cursor: pointer;
        font-size: 15px;
        font-weight: bold;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        transition: all 0.3s;
        width: 100%;
        display: block;
      }
      .automation-button:hover { transform: translateY(-3px); box-shadow: 0 6px 10px rgba(0,0,0,0.15); }
      .automation-button:active { transform: translateY(0); box-shadow: 0 3px 6px rgba(0,0,0,0.1); }
      .automation-button.stop {
        background: linear-gradient(145deg, #ff4d4f, #cc0000);
      }
      .automation-button.stop:hover { transform: translateY(-3px); box-shadow: 0 6px 10px rgba(0,0,0,0.15); }
      .automation-button.stop:active { transform: translateY(0); box-shadow: 0 3px 6px rgba(0,0,0,0.1); }
    `;
    injectStyle("automationStyles", css);
    const container = document.createElement("div");
    container.id = "automationContainer";
    container.className = "automation-container";
    container.style.display = "none";
    const startBtn = document.createElement("button");
    startBtn.textContent = "Start";
    startBtn.className = "automation-button";
    startBtn.addEventListener("click", startAuto);
    const stopBtn = document.createElement("button");
    stopBtn.textContent = "Stop";
    stopBtn.className = "automation-button stop";
    stopBtn.addEventListener("click", stopAuto);
    container.appendChild(startBtn);
    container.appendChild(stopBtn);
    document.body.appendChild(container);
    makeDraggable(container);
  };

  const createAutomationToggleButton = () => {
    if (qs("#toggleAutomationButton")) return;
    const btn = document.createElement("button");
    btn.id = "toggleAutomationButton";
    setStyles(btn, {
      position: "fixed", top: "100px", left: "20px", border: "none",
      borderRadius: "4px", padding: "10px 15px", cursor: "pointer", zIndex: "10000"
    });
    const updateStyle = () => {
      const active = sessionStorage.getItem("autoRunning") === "true";
      btn.style.backgroundColor = active ? "green" : "red";
      btn.style.color = "#fff";
      btn.textContent = active ? "Automatyzacja: Włączona" : "Automatyzacja: Wyłączona";
    };
    updateStyle();
    btn.addEventListener("click", () => {
      if (sessionStorage.getItem("autoRunning") === "true") {
        sessionStorage.setItem("autoRunning", "false");
        stopAuto();
      } else {
        sessionStorage.setItem("autoRunning", "true");
        startAuto();
      }
      updateStyle();
    });
    document.body.appendChild(btn);
  };

  // Automatyzacja – logika przeszukiwania i nawigacji
  const isAuto = () => sessionStorage.getItem("autoRunning") === "true";

  const getNextPage = () => {
    const nextLink = qs('a[rel="next"], a.next, .pagination-next a');
    if (nextLink?.href) return nextLink.href;
    try {
      const urlObj = new URL(location.href);
      if (urlObj.searchParams.has("page")) {
        let p = Number(urlObj.searchParams.get("page"));
        if (!isNaN(p)) {
          urlObj.searchParams.set("page", p + 1);
          return urlObj.toString();
        }
      }
    } catch (e) {
      console.error("URL error", e);
    }
    return null;
  };

  const runAuto = () => {
    if (!isAuto()) return console.log("Automatyzacja nieaktywna.");
    const els = qsa(".company-name.addax.addax-cs_hl_hit_company_name_click");
    if (els.length) {
      const target = qs(".col-lg-8.col-sm-7.align-self-center");
      target?.scrollIntoView({ behavior: "smooth" });
      setTimeout(() => { if (isAuto()) target?.click(); }, 500);
      const idx = Number(sessionStorage.getItem("clickedIndex")) || 0;
      if (idx >= els.length) {
        sessionStorage.removeItem("clickedIndex");
        return autoNext();
      }
      console.log(`Klikam ${idx + 1}/${els.length}`);
      els[idx].click();
      sessionStorage.setItem("clickedIndex", idx + 1);
    } else {
      const t = qs(".col-lg-8.col-sm-7.align-self-center");
      if (t) console.log("NIP:", t.textContent.match(cfg.reNIP) || [], "Telefony:", t.textContent.match(cfg.reTel) || []);
      setTimeout(() => { if (isAuto()) history.back(); }, 2000);
    }
  };

  const autoNext = () => {
    setTimeout(() => {
      if (!isAuto()) return;
      const np = getNextPage();
      if (np) {
        console.log("Przechodzę:", np);
        location.href = np;
      } else {
        console.log("Brak kolejnej strony.");
      }
    }, 2000);
  };

  const startAuto = () => {
    console.log("Automatyzacja uruchomiona.");
    sessionStorage.setItem("autoRunning", "true");
    runAuto();
  };

  const stopAuto = () => {
    console.log("Automatyzacja zatrzymana.");
    sessionStorage.setItem("autoRunning", "false");
    const toggle = qs("#toggleAutomationButton");
    if (toggle) {
      toggle.style.backgroundColor = "red";
      toggle.textContent = "Automatyzacja: Wyłączona";
    }
    const msg = document.createElement("div");
    setStyles(msg, {
      position: "fixed", top: "50%", left: "50%",
      transform: "translate(-50%, -50%)", backgroundColor: "green",
      color: "#fff", padding: "20px", borderRadius: "8px", zIndex: "10000",
      boxShadow: "0 0 10px rgba(0,0,0,0.5)"
    });
    msg.textContent = "Pobrano wszystkie wybrane wyniki.";
    document.body.appendChild(msg);
    setTimeout(() => { msg.remove(); }, 3000);
  };

  // Przycisk "Beta 0.2" – przeładowanie strony
  const createBetaButton = () => {
    const betaBtn = document.createElement("button");
    betaBtn.textContent = "Beta 0.2";
    setStyles(betaBtn, {
      position: "fixed", bottom: "20px", left: "50%",
      transform: "translateX(-50%)", backgroundColor: "#007bff",
      color: "#fff", border: "none", borderRadius: "4px", padding: "10px 15px",
      cursor: "pointer", zIndex: "9999"
    });
    betaBtn.addEventListener("click", () => {
      document.dispatchEvent(new KeyboardEvent("keydown", {
        key: "F5", code: "F5", keyCode: 116, which: 116, bubbles: true, cancelable: true
      }));
      location.reload();
    });
    document.body.appendChild(betaBtn);
  };

  // Główna funkcja – inicjalizacja skryptu
  const runScript = () => {
    if (
      location.href.includes("panoramafirm.pl") &&
      (document.title.trim().toLowerCase() === "wyniki dla fraz" ||
       (qs("h1") && qs("h1").textContent.trim().toLowerCase() === "wyniki dla fraz"))
    ) {
      console.log("Strona 'Wyniki dla fraz' na panoramafirm.pl – skrypt nie zostanie uruchomiony.");
      return;
    }
    if (localStorage.getItem(cfg.excelFlagKey) === "true") {
      localStorage.removeItem(cfg.excelFlagKey);
      saveData([]);
      updateUIPanel();
      return;
    }
    // Ukrywanie elementów map (Google Maps)
    qsa("img").forEach(img => {
      if (img.src?.includes("googleapis.com") || img.src?.includes("google.com/maps"))
        img.style.display = "none";
    });
    qsa("iframe").forEach(iframe => {
      if (iframe.src?.includes("google.com/maps"))
        iframe.style.display = "none";
    });
    qsa(".gm-style").forEach(el => el.style.display = "none");

    const data = collectData();
    if (data) {
      // Dodano warunki: jeśli jest więcej niż 2 e-maile lub nazwa firmy zaczyna się od "wyniki", pomijamy dodanie wiersza
      if (data.emaile.length > 2 || data.companyName.trim().toLowerCase().startsWith("wyniki")) {
        console.log("Wiersz pominięty: spełnia kryteria wykluczenia (więcej niż 2 e-maile lub nazwa firmy zaczyna się od 'Wyniki').");
      } else {
        const allData = getData();
        if (!isDuplicate(data, allData)) {
          allData.push(data);
          saveData(allData);
          console.log("Dane zapisane.");
        } else {
          console.log("Duplikat, pomijam.");
        }
      }
    }
    createUIPanel();
    updateUIPanel();
    if (localStorage.getItem(cfg.excelFlagKey) === "true") {
      localStorage.removeItem(cfg.excelFlagKey);
      generateExcel();
    }
    if (jsonData && jsonData.length >= 10 && !autoSearchTriggered) {
      autoUpdateSearchAndClick();
      autoSearchTriggered = true;
    }
  };

  // Funkcja aktualizująca wynik – pobiera dane z Excela i wykonuje wyszukiwanie
  function autoUpdateSearchAndClick() {
    const updateBtn = qs("#updateResultBtn");
    if (updateBtn) {
      updateBtn.disabled = true;
      updateBtn.textContent = "Przetwarzanie...";
    }
    // Pobieramy dane Excel zapisane w localStorage
    jsonData = localStorage.getItem("excelData") ? JSON.parse(localStorage.getItem("excelData")) : null;
    if (!jsonData || jsonData.length < 2) {
      if (updateBtn) {
        updateBtn.disabled = false;
        updateBtn.textContent = "Aktualizuj wynik";
      }
      return;
    }
    if (currentRowIndex >= jsonData.length) {
      console.log("Wszystkie wyniki zostały przetworzone.");
      const searchInput = document.getElementById("search-what");
      if (searchInput) searchInput.value = "Ostatnia strona - skrypt wyłączony";
      if (updateBtn) {
        updateBtn.disabled = false;
        updateBtn.textContent = "Aktualizuj wynik";
      }
      return;
    }
    const cellValue = jsonData[currentRowIndex][0];
    if (!cellValue) {
      if (updateBtn) {
        updateBtn.disabled = false;
        updateBtn.textContent = "Aktualizuj wynik";
      }
      return;
    }
    const searchInput = document.getElementById("search-what");
    if (!searchInput) {
      if (updateBtn) {
        updateBtn.disabled = false;
        updateBtn.textContent = "Aktualizuj wynik";
      }
      return;
    }
    searchInput.value = cellValue;
    searchInput.dispatchEvent(new Event("input", { bubbles: true }));
    searchInput.dispatchEvent(new Event("change", { bubbles: true }));
    const searchBtn = document.querySelector(
      ".btn.btn.search-btn.bg-primary.text-white.rounded-0.addax.addax-cs_hl_hit_company_name_click.labeled, .btn.btn.search-btn.bg-primary.text-white.rounded-0.addax.addax-cs_hl_hit_search.labeled"
    );
    if (searchBtn) searchBtn.click();
    console.log(`Wyszukano: ${cellValue}`);
    currentRowIndex++;
    const progressText = qs("#progressText");
    if (progressText) progressText.textContent = `Przetworzono ${currentRowIndex} z ${jsonData.length - 1} wyników`;
    setTimeout(() => {
      if (updateBtn) {
        updateBtn.disabled = false;
        updateBtn.textContent = "Aktualizuj wynik";
      }
    }, 1000);
  }

  // Inicjalizacja UI i uruchomienie skryptu
  createBetaButton();
  setupAutomationUI();
  createToggleButton();
  createAutomationToggleButton();
  runScript();
  if (isAuto()) runAuto();
})();
