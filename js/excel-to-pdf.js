document.addEventListener("DOMContentLoaded", function () {
  let selectedClient = "liberty";

  const excelFileInput = document.getElementById("excelFile");
  const excelFileName = document.getElementById("excelFileName");
  const certificateDateInput = document.getElementById("certificateDate");
  const previewBtn = document.getElementById("previewBtn");
  const generateBtn = document.getElementById("generateBtn");
  const statusText = document.getElementById("statusText");

  if (excelFileInput) {
    excelFileInput.addEventListener("change", function () {
      const file = excelFileInput.files && excelFileInput.files[0] ? excelFileInput.files[0] : null;
      if (excelFileName) excelFileName.innerText = file ? file.name : "No file chosen";
      if (statusText) statusText.innerText = file ? "Excel file selected. Enter certificate date and preview." : "No certificate file prepared yet";
    });
  }


  async function logCertificateRunSafe(data) {
    try {
      await fetch("/api/log-run", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(data)
      });
    } catch (err) {
      console.log("Log run skipped:", err);
    }
  }

  function getSelectedDeliveryOptions() {
    return {
      merged: document.getElementById("deliveryMerged").checked,
      separate: document.getElementById("deliverySeparate").checked,
      email: false
    };
  }

  function validateDeliveryOptions() {
    const options = getSelectedDeliveryOptions();

    if (!options.merged && !options.separate) {
      throw new Error("Please choose at least one active delivery method.");
    }

    return options;
  }

  function sanitizeFileName(value) {
    return String(value || "")
      .replace(/[\\\/:*?"<>|]/g, "")
      .replace(/\s+/g, " ")
      .trim();
  }

  function readDeliveryOptionLabels() {
    const labels = [];
    if (document.getElementById("deliveryMerged").checked) labels.push("Merged PDF");
    if (document.getElementById("deliverySeparate").checked) labels.push("Separate PDFs");
    labels.push("Email (Coming Soon)");
    return labels;
  }

  function updateTemplateNote() {
    const templateNote = document.getElementById("templateNote");

    if (selectedClient === "liberty") {
      templateNote.textContent = "Liberty template selected.";
      return;
    }

    if (selectedClient === "stanlib") {
      templateNote.textContent = "STANLIB template selected.";
      return;
    }

    if (selectedClient === "pm") {
      templateNote.textContent = "PM Skillshouse is visible in the parallel setup, but its certificate template is not loaded yet.";
      return;
    }

    if (selectedClient === "alexforbes") {
      templateNote.textContent = "AlexForbes is visible in the parallel setup, but its certificate template is not loaded yet.";
      return;
    }

    if (selectedClient === "sanlam") {
      templateNote.textContent = "Sanlam template is coming soon.";
    }
  }

  function formatCertificateDate(value) {
    if (!value) return "";
    const parts = value.split("-");
    if (parts.length !== 3) return value;
    return parts[2] + "-" + parts[1] + "-" + parts[0];
  }

  async function readExcelRows() {
    const certificateDate = formatCertificateDate(certificateDateInput.value);

    if (!excelFileInput.files || !excelFileInput.files[0]) {
      throw new Error("Please upload an Excel file.");
    }

    if (!certificateDate) {
      throw new Error("Please enter the certificate date.");
    }

    const excelBuffer = await excelFileInput.files[0].arrayBuffer();
    const workbook = XLSX.read(excelBuffer, { type: "array" });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

    if (!rows.length) {
      throw new Error("The Excel file has no data rows.");
    }

    const cleanedRows = rows
      .map(function (row) {
        return {
          fullName: String(row["Full Name"] || "").trim(),
          idNumber: String(row["ID Number"] || "").trim(),
          email: String(row["Email"] || "").trim()
        };
      })
      .filter(function (row) {
        return row.fullName && row.idNumber;
      });

    if (!cleanedRows.length) {
      throw new Error("No valid rows found. Please use headings exactly: Full Name, ID Number.");
    }

    const uniqueMap = new Map();

    cleanedRows.forEach(function (row) {
      const key = row.fullName + "||" + row.idNumber;
      if (!uniqueMap.has(key)) {
        uniqueMap.set(key, row);
      }
    });

    const uniqueRows = Array.from(uniqueMap.values());
    const duplicatesRemoved = cleanedRows.length - uniqueRows.length;
    const emailEligibleCount = uniqueRows.filter(function (row) {
      return !!row.email;
    }).length;
    const emailMissingCount = uniqueRows.length - emailEligibleCount;

    return {
      rows: uniqueRows,
      certificateDate: certificateDate,
      originalCount: cleanedRows.length,
      uniqueCount: uniqueRows.length,
      duplicatesRemoved: duplicatesRemoved,
      emailEligibleCount: emailEligibleCount,
      emailMissingCount: emailMissingCount
    };
  }

  async function loadTemplateBytes() {
    let filePath = "";

    if (selectedClient === "liberty") {
      filePath = "/templates/liberty-certificate.pdf";
    } else if (selectedClient === "stanlib") {
      filePath = "/templates/stanlib-certificate.pdf";
    } else if (selectedClient === "pm") {
      throw new Error("PM Skillshouse template is not loaded yet.");
    } else if (selectedClient === "alexforbes") {
      throw new Error("AlexForbes template is not loaded yet.");
    } else {
      throw new Error("Selected template is not available yet.");
    }

    const response = await fetch(filePath);

    if (!response.ok) {
      throw new Error("Certificate template could not be loaded.");
    }

    return await response.arrayBuffer();
  }

  function getTemplateSettings() {
    if (selectedClient === "stanlib") {
      return {
        nameYRatio: 0.50,
        idYRatio: 0.39,
        dateYRatio: 0.22,
        nameSize: 24,
        idSize: 18,
        dateSize: 15,
        rightMargin: 105
      };
    }

    return {
      nameYRatio: 0.515,
      idYRatio: 0.41,
      dateYRatio: 0.245,
      nameSize: 24,
      idSize: 20,
      dateSize: 16,
      rightMargin: 120
    };
  }

  async function buildSingleCertificatePdf(templateBytes, fullName, idNumber, certificateDate) {
    const pdfDoc = await PDFLib.PDFDocument.load(templateBytes);
    const page = pdfDoc.getPages()[0];

    const font = await pdfDoc.embedFont(PDFLib.StandardFonts.Helvetica);
    const boldFont = await pdfDoc.embedFont(PDFLib.StandardFonts.HelveticaBold);
    const textColor = PDFLib.rgb(0.2, 0.2, 0.2);
    const nameColor = PDFLib.rgb(0.1, 0.1, 0.1);

    const pageWidth = page.getWidth();
    const pageHeight = page.getHeight();
    const settings = getTemplateSettings();

    const nameWidth = boldFont.widthOfTextAtSize(fullName, settings.nameSize);
    const idWidth = font.widthOfTextAtSize(idNumber, settings.idSize);
    const dateWidth = font.widthOfTextAtSize(certificateDate, settings.dateSize);

    page.drawText(fullName, {
      x: (pageWidth - nameWidth) / 2,
      y: pageHeight * settings.nameYRatio,
      size: settings.nameSize,
      font: boldFont,
      color: nameColor
    });

    page.drawText(idNumber, {
      x: (pageWidth - idWidth) / 2,
      y: pageHeight * settings.idYRatio,
      size: settings.idSize,
      font: font,
      color: textColor
    });

    page.drawText(certificateDate, {
      x: pageWidth - dateWidth - settings.rightMargin,
      y: pageHeight * settings.dateYRatio,
      size: settings.dateSize,
      font: font,
      color: textColor
    });

    return pdfDoc;
  }

  function showPreviewSummary(result) {
    const previewBox = document.getElementById("previewBox");
    const previewCount = document.getElementById("previewCount");
    const previewList = document.getElementById("previewList");
    const selectedMethods = readDeliveryOptionLabels();

    let clientLabel = "Liberty";
    if (selectedClient === "pm") clientLabel = "PM Skillshouse";
    if (selectedClient === "stanlib") clientLabel = "STANLIB";
    if (selectedClient === "alexforbes") clientLabel = "AlexForbes";
    if (selectedClient === "sanlam") clientLabel = "Sanlam";

    previewCount.innerHTML =
      "Selected template: " + clientLabel + "<br>" +
      "Valid rows found: " + result.originalCount + "<br>" +
      "Unique records used: " + result.uniqueCount;

    if (result.duplicatesRemoved > 0) {
      previewCount.innerHTML +=
        "<br><span style='color:#b45309;font-weight:700;'>Duplicates detected and removed: " +
        result.duplicatesRemoved +
        "</span><br><span class='summary-text'>Only unique records will be used for certificate generation.</span>";
    }

    previewCount.innerHTML +=
      "<br><span class='summary-text'>Email-ready rows in file: " + result.emailEligibleCount +
      " | Missing email: " + result.emailMissingCount + "</span>";

    previewCount.innerHTML +=
      "<br><span class='summary-text'>Selected delivery: " + selectedMethods.join(", ") + "</span>";

    previewList.innerHTML = "";

    result.rows.slice(0, 5).forEach(function (row) {
      const li = document.createElement("li");
      li.textContent = row.fullName + " — " + row.idNumber;
      previewList.appendChild(li);
    });

    previewBox.style.display = "block";
  }

  async function previewFirstCertificate() {
    try {
      previewBtn.disabled = true;
      generateBtn.disabled = true;
      statusText.textContent = "Preparing preview...";

      const result = await readExcelRows();
      showPreviewSummary(result);

      const templateBytes = await loadTemplateBytes();
      const previewPdf = await buildSingleCertificatePdf(
        templateBytes,
        result.rows[0].fullName,
        result.rows[0].idNumber,
        result.certificateDate
      );

      const previewBytes = await previewPdf.save();
      const blob = new Blob([previewBytes], { type: "application/pdf" });
      const url = URL.createObjectURL(blob);

      const previewModal = document.getElementById("certificatePreviewModal");
      const previewFrame = document.getElementById("certificatePreviewFrame");

      if (previewFrame) previewFrame.src = url;
      if (previewModal) previewModal.style.display = "flex";

      statusText.textContent = "Preview ready.";
    } catch (error) {
      statusText.textContent = error.message || "Something went wrong while preparing the preview.";
    } finally {
      previewBtn.disabled = false;
      generateBtn.disabled = false;
    }
  }

  async function generateBulkCertificates() {
    try {
      previewBtn.disabled = true;
      generateBtn.disabled = true;
      statusText.textContent = "Preparing certificates...";

      const options = validateDeliveryOptions();
      const result = await readExcelRows();
      showPreviewSummary(result);

      const templateBytes = await loadTemplateBytes();

      let mergedPdf = null;
      let zip = null;

      if (options.merged) mergedPdf = await PDFLib.PDFDocument.create();
      if (options.separate) zip = new JSZip();

      for (let i = 0; i < result.rows.length; i++) {
        const row = result.rows[i];

        const singlePdf = await buildSingleCertificatePdf(
          templateBytes,
          row.fullName,
          row.idNumber,
          result.certificateDate
        );

        if (options.merged) {
          const copiedPages = await mergedPdf.copyPages(singlePdf, [0]);
          mergedPdf.addPage(copiedPages[0]);
        }

        if (options.separate) {
          const singleBytes = await singlePdf.save();
          const safeName = sanitizeFileName(row.fullName || "certificate");
          zip.file(safeName + ".pdf", singleBytes);
        }

        statusText.textContent =
          "Processing " + result.rows.length + " certificate(s)...\nCompleted: " +
          (i + 1) + " of " + result.rows.length;
      }

      if (options.merged) {
        const mergedBytes = await mergedPdf.save();
        const mergedBlob = new Blob([mergedBytes], { type: "application/pdf" });
        const mergedUrl = URL.createObjectURL(mergedBlob);

        const mergedLink = document.createElement("a");
        mergedLink.href = mergedUrl;
        const clientName = selectedClient === "stanlib" ? "STANLIB" : selectedClient === "liberty" ? "Liberty" : selectedClient;
mergedLink.download = clientName + " Certificates.pdf";
        mergedLink.click();
      }

      if (options.separate) {
        const zipBlob = await zip.generateAsync({ type: "blob" });
        const zipUrl = URL.createObjectURL(zipBlob);

        const zipLink = document.createElement("a");
        zipLink.href = zipUrl;
        const clientZipName = selectedClient === "stanlib" ? "STANLIB" : selectedClient === "liberty" ? "Liberty" : selectedClient;
zipLink.download = clientZipName + " Certificates ZIP.zip";
        zipLink.click();
      }

      const completedActions = [];
      if (options.merged) completedActions.push("merged PDF downloaded");
      if (options.separate) completedActions.push("ZIP of separate PDFs downloaded");

      await logCertificateRunSafe({
        client: selectedClient,
        certificate_count: result.rows.length,
        delivery_types: completedActions,
        source: "certificate-generator",
        notes: "Certificate generator run for " + selectedClient
      });

      statusText.textContent =
        "Done. " + result.rows.length + " certificate(s) processed successfully.\n" +
        "Output: " + completedActions.join(" and ") + ".";
    } catch (error) {
      statusText.textContent = error.message || "Something went wrong while processing the certificates.";
    } finally {
      previewBtn.disabled = false;
      generateBtn.disabled = false;
    }
  }

  document.querySelectorAll(".certificate-client-card").forEach(function (card) {
    card.addEventListener("click", function () {
      const client = card.getAttribute("data-client");

      if (client === "sanlam") {
        selectedClient = "sanlam";
      } else {
        selectedClient = client;
      }

      document.querySelectorAll(".certificate-client-card").forEach(function (item) {
        item.classList.remove("active");
      });

      card.classList.add("active");
      updateTemplateNote();
    });
  });


  const closeCertificatePreviewBtn = document.getElementById("closeCertificatePreviewBtn");
  if (closeCertificatePreviewBtn) {
    closeCertificatePreviewBtn.addEventListener("click", function () {
      const previewModal = document.getElementById("certificatePreviewModal");
      const previewFrame = document.getElementById("certificatePreviewFrame");
      if (previewFrame) previewFrame.src = "";
      if (previewModal) previewModal.style.display = "none";
    });
  }

  previewBtn.addEventListener("click", previewFirstCertificate);
  generateBtn.addEventListener("click", generateBulkCertificates);
});
