document.addEventListener("DOMContentLoaded", function () {
  let bookings = [];
  let activeBooking = null;

  const tableBody = document.getElementById("bookingTableBody");
  const searchInput = document.getElementById("bookingSearch");
  const statusFilter = document.getElementById("bookingStatusFilter");
  const sourceFilter = document.getElementById("bookingSourceFilter");

  function statusLabel(status) {
    return '<span class="booking-status-pill status-' + status + '">' + status + '</span>';
  }

  function sourceLabel(source) {
    return source === "client_excel" ? "Client Excel" : "Facilitator Request";
  }

  function renderBookings() {
    const search = (searchInput.value || "").toLowerCase();
    const status = statusFilter.value || "";
    const source = sourceFilter.value || "";

    const filtered = bookings.filter(function (booking) {
      const text = Object.values(booking).join(" ").toLowerCase();
      return (!search || text.includes(search)) &&
        (!status || booking.status === status) &&
        (!source || booking.source === source);
    });

    updateKpis();

    if (!filtered.length) {
      tableBody.innerHTML = '<tr><td colspan="8">No bookings loaded yet. Import a client Excel file to begin.</td></tr>';
      return;
    }

    tableBody.innerHTML = filtered.map(function (booking) {
      return `
        <tr>
          <td>${booking.date}</td>
          <td>${booking.time}</td>
          <td>${booking.client}<br><span class="summary-text">${sourceLabel(booking.source)}</span></td>
          <td>${booking.facilitator}</td>
          <td>${booking.company}</td>
          <td>${booking.province}</td>
          <td>${statusLabel(booking.status)}</td>
          <td>
            <div class="booking-actions">
              <button class="mini-dark-btn" type="button" onclick="openBookingDetails('${booking.id}')">View</button>
            </div>
          </td>
        </tr>
      `;
    }).join("");

    document.getElementById("bookingSummary").innerText = filtered.length + " booking(s) visible";
  }

  function updateKpis() {
    document.getElementById("pendingCount").innerText = bookings.filter(b => b.status === "pending").length;
    document.getElementById("approvedCount").innerText = bookings.filter(b => b.status === "approved").length;
    document.getElementById("allocatedCount").innerText = bookings.filter(b => b.status === "allocated").length;
    document.getElementById("completedCount").innerText = bookings.filter(b => b.status === "completed").length;
  }

  function whatsappMessage(booking) {
    return `Hi ${booking.facilitator === "Unallocated" ? "Facilitator" : booking.facilitator}, new ${booking.client} training booking:

Client: ${booking.client}
Date: ${booking.date}
Time: ${booking.time}
Venue: ${booking.venue}
Province: ${booking.province}
Participants: ${booking.participants}
Sessions: ${booking.sessions}
Language: ${booking.language}
Contact: ${booking.contact} - ${booking.contactNumber}

Please confirm if you accept this booking.`;
  }

  window.openBookingDetails = function (id) {
    activeBooking = bookings.find(function (booking) { return booking.id === id; });
    if (!activeBooking) return;

    document.getElementById("modalTitle").innerText = activeBooking.client + " Booking";
    document.getElementById("modalSubtitle").innerText = activeBooking.id + " • " + sourceLabel(activeBooking.source);

    const fields = [
      ["Status", activeBooking.status],
      ["Facilitator", activeBooking.facilitator],
      ["Training date", activeBooking.date],
      ["Start time", activeBooking.time],
      ["Participants", activeBooking.participants],
      ["Sessions", activeBooking.sessions],
      ["Community / Worksite", activeBooking.type],
      ["Province", activeBooking.province],
      ["Company / Community", activeBooking.company],
      ["Venue", activeBooking.venue],
      ["Contact person", activeBooking.contact],
      ["Contact number", activeBooking.contactNumber],
      ["Language", activeBooking.language],
      ["Notes", activeBooking.notes]
    ];

    document.getElementById("bookingDetailGrid").innerHTML = fields.map(function (field) {
      return `<div class="booking-detail-box"><span>${field[0]}</span><strong>${field[1]}</strong></div>`;
    }).join("");

    document.getElementById("whatsappPreview").innerText = whatsappMessage(activeBooking);
    document.getElementById("bookingModal").style.display = "flex";
  };

  document.getElementById("closeBookingModalBtn").addEventListener("click", function () {
    document.getElementById("bookingModal").style.display = "none";
  });

  document.getElementById("copyWhatsappBtn").addEventListener("click", async function () {
    if (!activeBooking) return;
    await navigator.clipboard.writeText(whatsappMessage(activeBooking));
    alert("WhatsApp message copied.");
  });

  document.getElementById("approveBookingBtn").addEventListener("click", function () {
    if (!activeBooking) return;
    activeBooking.status = "approved";
    renderBookings();
    window.openBookingDetails(activeBooking.id);
  });

  document.getElementById("allocateBookingBtn").addEventListener("click", function () {
    if (!activeBooking) return;
    activeBooking.status = "allocated";
    if (activeBooking.facilitator === "Unallocated") activeBooking.facilitator = "Julius Nthoba";
    renderBookings();
    window.openBookingDetails(activeBooking.id);
  });

  document.getElementById("declineBookingBtn").addEventListener("click", function () {
    if (!activeBooking) return;
    activeBooking.status = "declined";
    renderBookings();
    window.openBookingDetails(activeBooking.id);
  });

  searchInput.addEventListener("input", renderBookings);
  statusFilter.addEventListener("change", renderBookings);
  sourceFilter.addEventListener("change", renderBookings);




  const importBookingBtn = document.getElementById("importBookingBtn");
  const bookingImportFile = document.getElementById("bookingImportFile");

  function showBookingNotice(title, message, type) {
    let overlay = document.getElementById("bookingNoticeOverlay");

    if (!overlay) {
      overlay = document.createElement("div");
      overlay.id = "bookingNoticeOverlay";
      overlay.style.position = "fixed";
      overlay.style.inset = "0";
      overlay.style.background = "rgba(15,23,42,0.55)";
      overlay.style.zIndex = "10000";
      overlay.style.display = "none";
      overlay.style.alignItems = "center";
      overlay.style.justifyContent = "center";
      overlay.style.padding = "24px";

      overlay.innerHTML = `
        <div style="width:min(520px,94vw); background:#fff; border-radius:24px; box-shadow:0 30px 80px rgba(15,23,42,0.35); padding:26px;">
          <div id="bookingNoticeBadge" style="width:48px;height:48px;border-radius:16px;display:flex;align-items:center;justify-content:center;font-weight:900;margin-bottom:16px;">!</div>
          <h3 id="bookingNoticeTitle" style="margin:0 0 8px;font-size:22px;color:#111827;"></h3>
          <div id="bookingNoticeMessage" style="color:#6b7280;line-height:1.6;font-size:15px;"></div>
          <div style="display:flex;justify-content:flex-end;margin-top:22px;">
            <button id="bookingNoticeCloseBtn" type="button" style="width:auto;border:0;border-radius:14px;background:#0f172a;color:#fff;font-weight:800;padding:12px 18px;">Okay</button>
          </div>
        </div>
      `;

      document.body.appendChild(overlay);

      document.getElementById("bookingNoticeCloseBtn").addEventListener("click", function () {
        overlay.style.display = "none";
      });
    }

    const badge = document.getElementById("bookingNoticeBadge");
    const titleEl = document.getElementById("bookingNoticeTitle");
    const messageEl = document.getElementById("bookingNoticeMessage");

    if (type === "success") {
      badge.style.background = "#ecfdf5";
      badge.style.color = "#047857";
      badge.innerText = "✓";
    } else {
      badge.style.background = "#fff7ed";
      badge.style.color = "#c2410c";
      badge.innerText = "!";
    }

    titleEl.innerText = title;
    messageEl.innerText = message;
    overlay.style.display = "flex";
  }

  function normaliseKey(value) {
    return String(value || "").toLowerCase().replace(/[^a-z0-9]/g, "");
  }

  function cleanText(value) {
    return String(value === undefined || value === null ? "" : value).replace(/\s+/g, " ").trim();
  }

  function findValue(row, possibleKeys) {
    const keys = Object.keys(row || {});

    for (const wanted of possibleKeys) {
      const cleanWanted = normaliseKey(wanted);

      const found = keys.find(function (key) {
        const cleanKey = normaliseKey(key);
        return cleanKey === cleanWanted || cleanKey.includes(cleanWanted);
      });

      if (found && row[found] !== undefined && row[found] !== null && cleanText(row[found])) {
        return cleanText(row[found]);
      }
    }

    return "";
  }

  function excelDateToText(value) {
    if (!value) return "";
    if (String(value).toLowerCase().includes("e.g.")) return "";

    if (value instanceof Date) {
      return value.toISOString().slice(0, 10);
    }

    if (typeof value === "number") {
      const parsed = XLSX.SSF.parse_date_code(value);
      if (parsed) {
        const yyyy = parsed.y;
        const mm = String(parsed.m).padStart(2, "0");
        const dd = String(parsed.d).padStart(2, "0");
        return yyyy + "-" + mm + "-" + dd;
      }
    }

    if (value instanceof Date) {
      const yyyy = value.getFullYear();
      const mm = String(value.getMonth() + 1).padStart(2, "0");
      const dd = String(value.getDate()).padStart(2, "0");
      return yyyy + "-" + mm + "-" + dd;
    }

    return cleanText(value);
  }

  function guessClientFromFileName(name) {
    const lower = String(name || "").toLowerCase();

    if (lower.includes("stanlib")) return "STANLIB";
    if (lower.includes("liberty") || lower.includes("samwu") || lower.includes("pm skills")) return "Liberty";
    if (lower.includes("alex")) return "AlexForbes";

    return "Client";
  }

  function combineHeader(previous, current, index) {
    const prev = cleanText(previous[index]);
    let curr = cleanText(current[index]);

    if (index === 0 && curr.toLowerCase().includes("e.g.")) {
      curr = "Training Date";
    }

    if (!curr && prev) return prev;
    if (curr.toLowerCase().includes("e.g.") && prev) return prev;
    if (curr && prev && prev.toLowerCase().includes("training date")) return prev;
    if (curr && prev && prev.toLowerCase().includes("50 pax")) return curr;

    return curr || prev || "Column " + (index + 1);
  }

  function parseExcelBookings(rawRows) {
    let headerIndex = -1;

    for (let i = 0; i < rawRows.length; i++) {
      const rowText = rawRows[i].join(" ").toLowerCase();

      if (rowText.includes("date of workshop")) {
        headerIndex = i;
        break;
      }

      if (rowText.includes("start time") && rowText.includes("participants")) {
        headerIndex = i;
        break;
      }
    }

    if (headerIndex === -1) return [];

    const headerRow = rawRows[headerIndex] || [];

    return rawRows.slice(headerIndex + 1)
      .filter(function (row) {
        if (!row || !row.some(function (cell) { return cleanText(cell); })) return false;

        const first = cleanText(row[0]).toLowerCase();
        const second = cleanText(row[1]).toLowerCase();

        if (first.includes("e.g.")) return false;
        if (first.includes("training date")) return false;
        if (second.includes("start time")) return false;

        return true;
      })
      .map(function (row) {
        const obj = { __cells: row };
        headerRow.forEach(function (header, index) {
          const cleanHeader = cleanText(header);
          if (cleanHeader) obj[cleanHeader] = row[index];
        });
        return obj;
      });
  }


  function getByPosition(row, index) {
    const cells = row && row.__cells ? row.__cells : [];
    const value = cells[index];

    if (value instanceof Date) return value;

    return cleanText(value);
  }

  function buildBookingFromRow(row, fileName, index) {
    const client = guessClientFromFileName(fileName);
    const lowerFile = String(fileName || "").toLowerCase();

    let rawDate = "";
    let rawTime = "";
    let participants = "";
    let type = "";
    let deliveryStyle = "";
    let previousTraining = "";
    let facilitator = "";
    let facilitatorPhone = "";
    let representative = "";
    let department = "";
    let company = "";
    let venue = "";
    let province = "";
    let contact = "";
    let contactNumber = "";
    let language = "";
    let notes = "";

    if (lowerFile.includes("stanlib")) {
      rawDate = getByPosition(row, 0);
      company = getByPosition(row, 1);
      venue = getByPosition(row, 2);
      province = getByPosition(row, 3);
      rawTime = getByPosition(row, 4);
      contact = getByPosition(row, 5);
      contactNumber = getByPosition(row, 6);
      participants = getByPosition(row, 7);
      facilitator = getByPosition(row, 8);
      facilitatorPhone = getByPosition(row, 9);
      language = getByPosition(row, 10);
      notes = getByPosition(row, 11);
    } else {
      rawDate = getByPosition(row, 0);
      rawTime = getByPosition(row, 1);
      participants = getByPosition(row, 2);
      type = getByPosition(row, 3);
      deliveryStyle = getByPosition(row, 4);
      previousTraining = getByPosition(row, 5);
      facilitator = getByPosition(row, 6);
      facilitatorPhone = getByPosition(row, 7);
      representative = getByPosition(row, 8);
      department = getByPosition(row, 9);
      company = getByPosition(row, 10);
      venue = getByPosition(row, 11);
      province = getByPosition(row, 12);
      contact = getByPosition(row, 13);
      contactNumber = getByPosition(row, 14);
      language = getByPosition(row, 15);
    }

    return {
      id: "IMP-" + Date.now() + "-" + index,
      source: "client_excel",
      client: client,
      status: "approved",
      facilitator: facilitator || "Unallocated",
      facilitatorPhone: facilitatorPhone || "",
      date: excelDateToText(rawDate) || "Date not found",
      time: rawTime || "Time not found",
      sessions: "1",
      participants: participants || "Not stated",
      type: type || "Not stated",
      deliveryStyle: deliveryStyle || "Not stated",
      previousTraining: previousTraining || "Not stated",
      representative: representative || "Not stated",
      department: department || "Not stated",
      province: province || "Not stated",
      company: company || "Not stated",
      venue: venue || "Not stated",
      contact: contact || "Not stated",
      contactNumber: contactNumber || "Not stated",
      language: language || "Not stated",
      notes: notes || "Imported from " + fileName
    };
  }

  async function importBookingExcel(file) {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    const rows = parseExcelBookings(rawRows);

    if (!rows.length) {
      showBookingNotice("Import not completed", "The file was read, but no usable booking rows were detected.", "warning");
      return;
    }

    const imported = rows
      .map(function (row, index) {
        return buildBookingFromRow(row, file.name, index + 1);
      })
      .filter(function (booking) {
        return booking.date !== "Date not found" || booking.venue !== "Not stated" || booking.company !== "Not stated";
      });

    if (!imported.length) {
      showBookingNotice("Import not completed", "The file was read, but no usable booking rows were detected.", "warning");
      return;
    }

    bookings = imported.concat(bookings);
    renderBookings();

    showBookingNotice("Booking import complete", imported.length + " booking(s) imported for admin review.", "success");
  }

  if (importBookingBtn && bookingImportFile) {
    importBookingBtn.addEventListener("click", function () {
      bookingImportFile.click();
    });

    bookingImportFile.addEventListener("change", async function () {
      const file = bookingImportFile.files && bookingImportFile.files[0] ? bookingImportFile.files[0] : null;
      if (!file) return;

      try {
        await importBookingExcel(file);
      } catch (error) {
        showBookingNotice("Import failed", error && error.message ? error.message : "Could not import booking file.", "warning");
      } finally {
        bookingImportFile.value = "";
      }
    });
  }

  document.getElementById("newBookingBtn").addEventListener("click", function () {
    alert("Manual booking form will be added next.");
  });

  renderBookings();
});
