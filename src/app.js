/* ============================================
   Sınav Salon Dağıtım - Main Application
   Offline-capable, no external dependencies
   ============================================ */

(function () {
  'use strict';

  // ===== STATE =====
  const state = {
    step: 1,
    students: [],
    rooms: [],
    distribution: null,
    editingRoomId: null,
    studentPage: 1,
    studentsPerPage: 25
  };

  // ===== DOM REFS =====
  let DOM = {};

  // ===== INIT =====
  function init() {
    cacheDOMRefs();
    setupEventListeners();
    setupDragAndDrop('student-upload-area', handleStudentUpload);
    setupDragAndDrop('room-upload-area', handleRoomUpload);
    updateUI();
  }

  function cacheDOMRefs() {
    DOM.steps = {
      1: document.getElementById('step-1'),
      2: document.getElementById('step-2'),
      3: document.getElementById('step-3')
    };
    DOM.indicators = {
      1: document.getElementById('step-indicator-1'),
      2: document.getElementById('step-indicator-2'),
      3: document.getElementById('step-indicator-3')
    };
    DOM.connectors = {
      1: document.getElementById('connector-1'),
      2: document.getElementById('connector-2')
    };
    DOM.studentCount = document.getElementById('student-count');
    DOM.roomCount = document.getElementById('room-count');
    DOM.totalCapacity = document.getElementById('total-capacity');
    DOM.studentPreview = document.getElementById('student-preview');
    DOM.studentUploadArea = document.getElementById('student-upload-area');
    DOM.toastContainer = document.getElementById('toast-container');
  }

  // ===== TOAST NOTIFICATIONS =====
  function showToast(message, type) {
    type = type || 'info';
    var icons = {
      success: '✅',
      error: '❌',
      warning: '⚠️',
      info: 'ℹ️'
    };
    var toast = document.createElement('div');
    toast.className = 'toast toast-' + type;
    toast.innerHTML = '<span>' + (icons[type] || icons.info) + '</span><span>' + message + '</span>';
    DOM.toastContainer.appendChild(toast);

    setTimeout(function () {
      toast.classList.add('toast-removing');
      setTimeout(function () { toast.remove(); }, 300);
    }, 3500);
  }

  // ===== DRAG & DROP =====
  function setupDragAndDrop(areaId, callback) {
    var area = document.getElementById(areaId);
    if (!area) return;

    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(function (evt) {
      area.addEventListener(evt, function (e) {
        e.preventDefault();
        e.stopPropagation();
      }, false);
    });

    ['dragenter', 'dragover'].forEach(function (evt) {
      area.addEventListener(evt, function () { area.classList.add('drag-over'); }, false);
    });

    ['dragleave', 'drop'].forEach(function (evt) {
      area.addEventListener(evt, function () { area.classList.remove('drag-over'); }, false);
    });

    area.addEventListener('drop', function (e) {
      var files = e.dataTransfer.files;
      if (files.length > 0) callback(files[0]);
    }, false);
  }

  // ===== UI UPDATE =====
  function updateUI() {
    // Steps visibility & indicators
    for (var key = 1; key <= 3; key++) {
      if (key === state.step) {
        DOM.steps[key].classList.remove('hidden');
        DOM.steps[key].classList.add('animate-in');
        DOM.indicators[key].classList.add('active');
        DOM.indicators[key].classList.remove('completed');
      } else {
        DOM.steps[key].classList.add('hidden');
        DOM.indicators[key].classList.remove('active');
      }
      if (key < state.step) {
        DOM.indicators[key].classList.add('completed');
      } else if (key > state.step) {
        DOM.indicators[key].classList.remove('completed');
      }
    }

    // Connectors
    for (var c = 1; c <= 2; c++) {
      DOM.connectors[c].classList.remove('completed', 'active');
      if (c < state.step) DOM.connectors[c].classList.add('completed');
      else if (c === state.step) DOM.connectors[c].classList.add('active');
    }

    // Counts
    DOM.studentCount.textContent = state.students.length;
    DOM.roomCount.textContent = state.rooms.length;
    var totalCap = state.rooms.reduce(function (sum, r) { return sum + r.capacity; }, 0);
    DOM.totalCapacity.textContent = totalCap;
  }

  function setStep(step) {
    state.step = step;
    updateUI();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  // ===== EXCEL PARSING =====
  function parseExcelFile(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function (e) {
        try {
          var data = new Uint8Array(e.target.result);
          var workbook = XLSX.read(data, { type: 'array' });
          var sheetName = workbook.SheetNames[0];
          var worksheet = workbook.Sheets[sheetName];
          var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
          resolve(jsonData);
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = function (err) { reject(err); };
      reader.readAsArrayBuffer(file);
    });
  }

  // ===== STUDENT HANDLING =====
  function handleStudentUpload(file) {
    parseExcelFile(file).then(function (data) {
      if (!data || data.length === 0) {
        showToast('Dosya boş. Lütfen geçerli bir dosya yükleyin.', 'error');
        return;
      }
      if (data.length === 1) {
        showToast('Dosyada öğrenci kaydı bulunamadı.', 'error');
        return;
      }

      var rows = data.slice(1);
      state.students = rows.map(function (row) {
        var name = (row[0] != null) ? String(row[0]) : '';
        var surname = (row[1] != null) ? String(row[1]) : '';
        var fullName = (name + ' ' + surname).trim();

        return {
          id: (row[3] != null) ? String(row[3]) : '',
          tc: (row[4] != null) ? String(row[4]) : '',
          name: fullName,
          phone: (row[2] != null) ? String(row[2]) : '',
          classRef: (row[5] != null) ? String(row[5]) : '',
          department: (row[6] != null) ? String(row[6]) : ''
        };
      }).filter(function (s) { return s.name.length > 0; });
      state.studentPage = 1;

      renderStudentTable();
      updateUI();
      DOM.studentPreview.classList.remove('hidden');
      DOM.studentUploadArea.classList.add('hidden');
      showToast(state.students.length + ' öğrenci başarıyla yüklendi.', 'success');

    }).catch(function (err) {
      console.error(err);
      showToast('Dosya okuma hatası: ' + err.message, 'error');
    });
  }

  function renderStudentTable() {
    var thead = document.querySelector('#student-table thead');
    var tbody = document.querySelector('#student-table tbody');

    thead.innerHTML =
      '<tr>' +
      '<th>#</th>' +
      '<th>No</th>' +
      '<th>TC No</th>' +
      '<th>Ad Soyad</th>' +
      '<th>Sınıf</th>' +
      '<th>Bölüm</th>' +
      '<th>Telefon</th>' +
      '</tr>';

    tbody.innerHTML = '';

    var perPage = state.studentsPerPage;
    var page = state.studentPage;
    var totalPages = Math.ceil(state.students.length / perPage);
    if (page > totalPages) page = totalPages;
    if (page < 1) page = 1;
    state.studentPage = page;

    var startIdx = (page - 1) * perPage;
    var endIdx = Math.min(startIdx + perPage, state.students.length);
    var list = state.students.slice(startIdx, endIdx);

    list.forEach(function (s, i) {
      var tr = document.createElement('tr');
      tr.innerHTML =
        '<td>' + (startIdx + i + 1) + '</td>' +
        '<td>' + escapeHtml(s.id) + '</td>' +
        '<td>' + escapeHtml(s.tc) + '</td>' +
        '<td>' + escapeHtml(s.name) + '</td>' +
        '<td>' + escapeHtml(s.classRef) + '</td>' +
        '<td>' + escapeHtml(s.department) + '</td>' +
        '<td>' + escapeHtml(s.phone) + '</td>';
      tbody.appendChild(tr);
    });

    // Render pagination
    renderStudentPagination(totalPages);
  }

  function renderStudentPagination(totalPages) {
    var existingPag = document.getElementById('student-pagination');
    if (existingPag) existingPag.remove();

    if (totalPages <= 1) return;

    var container = document.createElement('div');
    container.id = 'student-pagination';
    container.className = 'pagination';

    // Prev button
    var prevBtn = document.createElement('button');
    prevBtn.className = 'pagination-btn' + (state.studentPage <= 1 ? ' disabled' : '');
    prevBtn.innerHTML = '&laquo;';
    prevBtn.disabled = state.studentPage <= 1;
    prevBtn.addEventListener('click', function () {
      if (state.studentPage > 1) { state.studentPage--; renderStudentTable(); }
    });
    container.appendChild(prevBtn);

    // Page numbers
    var startPage = Math.max(1, state.studentPage - 2);
    var endPage = Math.min(totalPages, startPage + 4);
    if (endPage - startPage < 4) startPage = Math.max(1, endPage - 4);

    if (startPage > 1) {
      container.appendChild(createPageBtn(1));
      if (startPage > 2) {
        var dots = document.createElement('span');
        dots.className = 'pagination-dots';
        dots.textContent = '...';
        container.appendChild(dots);
      }
    }

    for (var p = startPage; p <= endPage; p++) {
      container.appendChild(createPageBtn(p));
    }

    if (endPage < totalPages) {
      if (endPage < totalPages - 1) {
        var dots2 = document.createElement('span');
        dots2.className = 'pagination-dots';
        dots2.textContent = '...';
        container.appendChild(dots2);
      }
      container.appendChild(createPageBtn(totalPages));
    }

    // Next button
    var nextBtn = document.createElement('button');
    nextBtn.className = 'pagination-btn' + (state.studentPage >= totalPages ? ' disabled' : '');
    nextBtn.innerHTML = '&raquo;';
    nextBtn.disabled = state.studentPage >= totalPages;
    nextBtn.addEventListener('click', function () {
      if (state.studentPage < totalPages) { state.studentPage++; renderStudentTable(); }
    });
    container.appendChild(nextBtn);

    // Info text
    var info = document.createElement('span');
    info.className = 'pagination-info';
    info.textContent = state.students.length + ' öğrenci, Sayfa ' + state.studentPage + '/' + totalPages;
    container.appendChild(info);

    // Insert after table
    var tableContainer = document.querySelector('#student-table').closest('.data-table-container');
    tableContainer.parentNode.insertBefore(container, tableContainer.nextSibling);
  }

  function createPageBtn(pageNum) {
    var btn = document.createElement('button');
    btn.className = 'pagination-btn' + (pageNum === state.studentPage ? ' active' : '');
    btn.textContent = pageNum;
    btn.addEventListener('click', function () {
      state.studentPage = pageNum;
      renderStudentTable();
    });
    return btn;
  }

  // ===== ROOM HANDLING =====
  function handleRoomUpload(file) {
    parseExcelFile(file).then(function (data) {
      if (!data || data.length < 2) {
        showToast('Dosyada salon bilgisi bulunamadı.', 'error');
        return;
      }

      var rows = data.slice(1);
      var newRooms = rows.map(function (row, index) {
        return {
          id: Date.now() + index,
          name: row[0] || ('Salon ' + (index + 1)),
          capacity: parseInt(row[1]) || 20,
          priority: parseInt(row[2]) || 999
        };
      }).filter(function (r) { return r.capacity > 0; });

      state.rooms = state.rooms.concat(newRooms);
      state.rooms.sort(function (a, b) { return a.priority - b.priority; });
      updateUI();
      renderRoomTable();
      showToast(newRooms.length + ' salon başarıyla eklendi.', 'success');

    }).catch(function (err) {
      console.error(err);
      showToast('Dosya okuma hatası: ' + err.message, 'error');
    });
  }

  function renderRoomTable() {
    var tbody = document.querySelector('#room-table tbody');
    var nameInput = document.getElementById('manual-room-name');
    var capInput = document.getElementById('manual-room-capacity');
    var priorityInput = document.getElementById('manual-room-priority');
    var btnAdd = document.getElementById('btn-add-room');

    tbody.innerHTML = '';
    state.rooms.sort(function (a, b) { return a.priority - b.priority; });

    state.rooms.forEach(function (r, index) {
      var tr = document.createElement('tr');
      if (r.id === state.editingRoomId) tr.classList.add('editing');

      tr.innerHTML =
        '<td>' + escapeHtml(r.name) + '</td>' +
        '<td>' + r.capacity + '</td>' +
        '<td>' + (r.priority === 999 ? '—' : r.priority) + '</td>' +
        '<td><button class="btn-delete" data-index="' + index + '"><svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg> Sil</button></td>';

      // Double click to edit
      tr.addEventListener('dblclick', function () {
        state.editingRoomId = r.id;
        nameInput.value = r.name;
        capInput.value = r.capacity;
        priorityInput.value = r.priority === 999 ? '' : r.priority;
        btnAdd.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="20 6 9 17 4 12"/></svg>';
        btnAdd.classList.remove('btn-accent');
        btnAdd.classList.add('btn-warning');
        renderRoomTable();
        nameInput.focus();
      });

      tbody.appendChild(tr);
    });

    // Delete handlers
    document.querySelectorAll('.btn-delete').forEach(function (btn) {
      btn.addEventListener('click', function (e) {
        var idx = parseInt(e.currentTarget.getAttribute('data-index'));
        var room = state.rooms[idx];
        if (room && state.editingRoomId === room.id) {
          cancelEditMode();
        }
        state.rooms.splice(idx, 1);
        updateUI();
        renderRoomTable();
        showToast('Salon silindi.', 'warning');
      });
    });
  }

  function cancelEditMode() {
    state.editingRoomId = null;
    var nameInput = document.getElementById('manual-room-name');
    var capInput = document.getElementById('manual-room-capacity');
    var priorityInput = document.getElementById('manual-room-priority');
    var btnAdd = document.getElementById('btn-add-room');

    nameInput.value = '';
    capInput.value = '';
    priorityInput.value = '';
    btnAdd.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>';
    btnAdd.classList.remove('btn-warning');
    btnAdd.classList.add('btn-accent');
  }

  // ===== DISTRIBUTION =====
  function distributeStudents() {
    var studentsToDistribute = state.students.slice();
    shuffleArray(studentsToDistribute);

    var result = {};
    state.rooms.forEach(function (r) { result[r.id] = []; });

    var overflow = [];
    var currentRoomIndex = 0;

    studentsToDistribute.forEach(function (student) {
      var placed = false;

      while (currentRoomIndex < state.rooms.length) {
        var room = state.rooms[currentRoomIndex];
        if (result[room.id].length < room.capacity) {
          result[room.id].push(student);
          placed = true;
          break;
        } else {
          currentRoomIndex++;
        }
      }

      if (!placed) overflow.push(student);
    });

    state.distribution = { results: result, overflow: overflow };
    renderDistributionResults();
  }

  function shuffleArray(array) {
    for (var i = array.length - 1; i > 0; i--) {
      var j = Math.floor(Math.random() * (i + 1));
      var temp = array[i];
      array[i] = array[j];
      array[j] = temp;
    }
  }

  function renderDistributionResults() {
    var summaryEl = document.getElementById('distribution-summary');
    var resultsEl = document.getElementById('distribution-results');
    var totalDistributed = state.students.length - state.distribution.overflow.length;
    var hasOverflow = state.distribution.overflow.length > 0;

    // Summary cards
    summaryEl.innerHTML =
      '<div class="summary-grid">' +
      '<div class="summary-card summary-total">' +
      '<div class="summary-value">' + state.students.length + '</div>' +
      '<div class="summary-label">Toplam Öğrenci</div>' +
      '</div>' +
      '<div class="summary-card summary-placed">' +
      '<div class="summary-value">' + totalDistributed + '</div>' +
      '<div class="summary-label">Yerleşen</div>' +
      '</div>' +
      '<div class="summary-card ' + (hasOverflow ? 'summary-overflow' : 'summary-overflow no-overflow') + '">' +
      '<div class="summary-value">' + state.distribution.overflow.length + '</div>' +
      '<div class="summary-label">' + (hasOverflow ? 'Açıkta Kalan' : 'Herkes Yerleşti!') + '</div>' +
      '</div>' +
      '</div>' +
      (hasOverflow ?
        '<div style="margin-bottom: 1.5rem;"><button id="btn-export-overflow" class="btn btn-danger btn-sm">⚠️ Açıkta Kalanları İndir (Excel)</button></div>'
        : '');

    // Overflow export button
    var btnOverflow = document.getElementById('btn-export-overflow');
    if (btnOverflow) {
      btnOverflow.addEventListener('click', function () {
        exportOverflowToExcel(state.distribution.overflow);
      });
    }

    // Room result cards
    var html = '';

    state.rooms.forEach(function (room) {
      var students = state.distribution.results[room.id];
      html +=
        '<div class="room-result-card">' +
        '<div class="room-result-header">' +
        '<h3>' + escapeHtml(room.name) + '</h3>' +
        '<span class="capacity-badge">' + students.length + ' / ' + room.capacity + '</span>' +
        '</div>' +
        '<div class="data-table-container">' +
        '<table>' +
        '<thead><tr><th>Sıra</th><th>No</th><th>TC No</th><th>Ad Soyad</th><th>Sınıf</th><th>Bölüm</th><th>Telefon</th></tr></thead>' +
        '<tbody>' +
        (students.length === 0 ?
          '<tr><td colspan="7" style="text-align:center; color:var(--text-muted);">Öğrenci yok</td></tr>' :
          students.map(function (s, i) {
            return '<tr>' +
              '<td>' + (i + 1) + '</td>' +
              '<td>' + escapeHtml(String(s.id)) + '</td>' +
              '<td>' + escapeHtml(String(s.tc)) + '</td>' +
              '<td>' + escapeHtml(s.name) + '</td>' +
              '<td>' + escapeHtml(String(s.classRef)) + '</td>' +
              '<td>' + escapeHtml(String(s.department)) + '</td>' +
              '<td>' + escapeHtml(String(s.phone)) + '</td>' +
              '</tr>';
          }).join('')) +
        '</tbody>' +
        '</table>' +
        '</div>' +
        '</div>';
    });

    // Overflow section
    if (hasOverflow) {
      html +=
        '<div class="room-result-card overflow-result-card">' +
        '<div class="room-result-header">' +
        '<h3>⚠️ Açıkta Kalan Öğrenciler</h3>' +
        '<span class="capacity-badge">' + state.distribution.overflow.length + ' öğrenci</span>' +
        '</div>' +
        '<div class="data-table-container">' +
        '<table>' +
        '<thead><tr><th>No</th><th>TC No</th><th>Ad Soyad</th><th>Sınıf</th><th>Bölüm</th><th>Telefon</th></tr></thead>' +
        '<tbody>' +
        state.distribution.overflow.map(function (s) {
          return '<tr>' +
            '<td>' + escapeHtml(String(s.id)) + '</td>' +
            '<td>' + escapeHtml(String(s.tc)) + '</td>' +
            '<td>' + escapeHtml(s.name) + '</td>' +
            '<td>' + escapeHtml(String(s.classRef)) + '</td>' +
            '<td>' + escapeHtml(String(s.department)) + '</td>' +
            '<td>' + escapeHtml(String(s.phone)) + '</td>' +
            '</tr>';
        }).join('') +
        '</tbody>' +
        '</table>' +
        '</div>' +
        '</div>';
    }

    resultsEl.innerHTML = html;
  }

  // ===== EXCEL EXPORT =====
  function exportToExcel(data, rooms) {
    var wb = XLSX.utils.book_new();

    // Overview sheet
    var overviewData = [];
    rooms.forEach(function (room) {
      var students = data.results[room.id] || [];
      students.forEach(function (s, i) {
        overviewData.push({
          'Salon Adı': room.name,
          'Sıra No': i + 1,
          'Öğrenci No': s.id,
          'TC No': s.tc,
          'Ad Soyad': s.name,
          'Sınıf': s.classRef,
          'Bölüm': s.department,
          'Telefon': s.phone
        });
      });
    });

    if (data.overflow && data.overflow.length > 0) {
      data.overflow.forEach(function (s) {
        overviewData.push({
          'Salon Adı': 'YERLEŞEMEDİ',
          'Sıra No': '-',
          'Öğrenci No': s.id,
          'TC No': s.tc,
          'Ad Soyad': s.name,
          'Sınıf': s.classRef,
          'Bölüm': s.department,
          'Telefon': s.phone
        });
      });
    }

    var wsOverview = XLSX.utils.json_to_sheet(overviewData);
    XLSX.utils.book_append_sheet(wb, wsOverview, 'Genel Liste');

    // Per-room sheets
    rooms.forEach(function (room) {
      var students = data.results[room.id] || [];
      if (students.length > 0) {
        var roomData = students.map(function (s, i) {
          return {
            'Sıra No': i + 1,
            'Öğrenci No': s.id,
            'TC No': s.tc,
            'Ad Soyad': s.name,
            'Sınıf': s.classRef,
            'Bölüm': s.department,
            'Telefon': s.phone
          };
        });
        var wsRoom = XLSX.utils.json_to_sheet(roomData);

        var sheetName = room.name.replace(/[\\/?*[\]:]/g, ' ').trim();
        if (sheetName.length > 31) sheetName = sheetName.substring(0, 31);
        if (!sheetName) sheetName = 'Salon ' + room.id;

        var uniqueName = sheetName;
        var counter = 1;
        while (wb.SheetNames.indexOf(uniqueName) !== -1) {
          uniqueName = sheetName.substring(0, 28) + '(' + counter + ')';
          counter++;
        }

        XLSX.utils.book_append_sheet(wb, wsRoom, uniqueName);
      }
    });

    XLSX.writeFile(wb, 'sinav_dagitim_sonuclari.xlsx');
    showToast('Excel dosyası indirildi.', 'success');
  }

  function exportRoomsToExcel(rooms) {
    var wb = XLSX.utils.book_new();
    var sorted = rooms.slice().sort(function (a, b) { return (a.priority || 999) - (b.priority || 999); });
    var exportData = sorted.map(function (r) {
      return {
        'Salon Adı': r.name,
        'Kapasite': r.capacity,
        'Öncelik': r.priority === 999 ? '' : r.priority
      };
    });
    var ws = XLSX.utils.json_to_sheet(exportData);
    ws['!cols'] = [{ wch: 20 }, { wch: 10 }, { wch: 10 }];
    XLSX.utils.book_append_sheet(wb, ws, 'Salonlar');
    XLSX.writeFile(wb, 'salon_listesi.xlsx');
    showToast('Salon listesi indirildi.', 'success');
  }

  function exportOverflowToExcel(overflowData) {
    var wb = XLSX.utils.book_new();
    var exportData = overflowData.map(function (s) {
      return {
        'Öğrenci No': s.id,
        'TC No': s.tc,
        'Ad Soyad': s.name,
        'Sınıf': s.classRef,
        'Bölüm': s.department,
        'Telefon': s.phone
      };
    });
    var ws = XLSX.utils.json_to_sheet(exportData);
    ws['!cols'] = [{ wch: 15 }, { wch: 20 }, { wch: 30 }, { wch: 10 }, { wch: 20 }, { wch: 15 }];
    XLSX.utils.book_append_sheet(wb, ws, 'Açıkta Kalanlar');
    XLSX.writeFile(wb, 'acikta_kalanlar.xlsx');
    showToast('Açıkta kalanlar listesi indirildi.', 'success');
  }

  // ===== PDF EXPORT =====
  function generatePDF(distributionData, rooms) {
    var jspdf = window.jspdf;
    var doc = new jspdf.jsPDF();

    var sortedRooms = rooms.slice().sort(function (a, b) { return (a.priority || 999) - (b.priority || 999); });
    var isFirstPage = true;

    sortedRooms.forEach(function (room) {
      var students = distributionData.results[room.id] || [];
      if (students.length === 0) return;

      if (!isFirstPage) {
        doc.addPage();
      } else {
        isFirstPage = false;
      }

      doc.setFontSize(16);
      doc.text('Salon: ' + transliterate(room.name), 14, 20);
      doc.setFontSize(10);
      doc.text('Kapasite: ' + students.length + ' / ' + room.capacity, 14, 28);

      var tableData = students.map(function (s, i) {
        return [
          i + 1,
          transliterate(String(s.id)),
          transliterate(String(s.tc)),
          transliterate(s.name),
          transliterate(String(s.classRef)),
          transliterate(String(s.department)),
          transliterate(String(s.phone))
        ];
      });

      doc.autoTable({
        startY: 35,
        head: [['SIRA', 'NO', 'TC', 'AD SOYAD', 'SINIF', 'BOLUM', 'TEL']],
        body: tableData,
        theme: 'grid',
        headStyles: { fillColor: [79, 70, 229] },
        styles: { fontSize: 9, cellPadding: 2 },
        columnStyles: {
          0: { cellWidth: 15 },
          1: { cellWidth: 20 },
          2: { cellWidth: 25 },
          6: { cellWidth: 25 }
        }
      });
    });

    // Overflow page
    var overflow = distributionData.overflow || [];
    if (overflow.length > 0) {
      if (!isFirstPage) doc.addPage();
      doc.setFontSize(16);
      doc.setTextColor(220, 38, 38);
      doc.text('Acikta Kalanlar Listesi', 14, 20);
      doc.setTextColor(0, 0, 0);

      var tableData = overflow.map(function (s, i) {
        return [
          i + 1,
          transliterate(String(s.id)),
          transliterate(String(s.tc)),
          transliterate(s.name),
          transliterate(String(s.classRef)),
          transliterate(String(s.department)),
          transliterate(String(s.phone))
        ];
      });

      doc.autoTable({
        startY: 30,
        head: [['SIRA', 'NO', 'TC', 'AD SOYAD', 'SINIF', 'BOLUM', 'TEL']],
        body: tableData,
        theme: 'striped',
        headStyles: { fillColor: [220, 38, 38] },
        styles: { fontSize: 9 }
      });
    }

    doc.save('sinav_dagitim_raporu.pdf');
    showToast('PDF raporu indirildi.', 'success');
  }

  // Turkish char transliteration for PDF (jsPDF default font limitation)
  function transliterate(text) {
    if (!text) return '';
    var map = {
      'ğ': 'g', 'Ğ': 'G',
      'ş': 's', 'Ş': 'S',
      'ı': 'i', 'İ': 'I',
      'ç': 'c', 'Ç': 'C',
      'ö': 'o', 'Ö': 'O',
      'ü': 'u', 'Ü': 'U'
    };
    return text.replace(/[ğĞşŞıİçÇöÖüÜ]/g, function (ch) {
      return map[ch] || ch;
    });
  }

  // ===== ESCAPE HTML =====
  function escapeHtml(text) {
    var div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  // ===== EVENT LISTENERS =====
  function setupEventListeners() {
    // Student file input
    document.getElementById('student-file-input').addEventListener('change', function (e) {
      if (e.target.files.length > 0) handleStudentUpload(e.target.files[0]);
    });

    // Reload students
    document.getElementById('btn-reload-students').addEventListener('click', function () {
      state.students = [];
      DOM.studentPreview.classList.add('hidden');
      DOM.studentUploadArea.classList.remove('hidden');
      document.getElementById('student-file-input').value = '';
      updateUI();
      showToast('Öğrenci listesi temizlendi. Yeni dosya yükleyebilirsiniz.', 'info');
    });

    // Nav: Step 1 → 2
    document.getElementById('btn-to-step-2').addEventListener('click', function () {
      if (state.students.length === 0) {
        showToast('Lütfen önce öğrenci listesi yükleyin.', 'warning');
        return;
      }
      setStep(2);
    });

    // Nav: Step 2 → 1
    document.getElementById('btn-back-to-step-1').addEventListener('click', function () { setStep(1); });

    // Room file input
    document.getElementById('room-file-input').addEventListener('change', function (e) {
      if (e.target.files.length > 0) handleRoomUpload(e.target.files[0]);
    });

    // Manual room add
    document.getElementById('btn-add-room').addEventListener('click', function () {
      var nameInput = document.getElementById('manual-room-name');
      var capInput = document.getElementById('manual-room-capacity');
      var priorityInput = document.getElementById('manual-room-priority');
      var btnAdd = document.getElementById('btn-add-room');

      var name = nameInput.value.trim();
      var cap = parseInt(capInput.value);
      var priority = parseInt(priorityInput.value) || 999;

      if (!name || isNaN(cap) || cap <= 0) {
        showToast('Lütfen geçerli bir salon adı ve kapasite giriniz.', 'warning');
        return;
      }

      // Check duplicate priority
      if (priority !== 999) {
        var duplicate = state.rooms.find(function (r) { return r.priority === priority && r.id !== state.editingRoomId; });
        if (duplicate) {
          if (!confirm(priority + ' öncelik sırası zaten "' + duplicate.name + '" salonunda kullanılıyor. Devam edilsin mi?')) {
            return;
          }
        }
      }

      if (state.editingRoomId) {
        var roomIndex = state.rooms.findIndex(function (r) { return r.id === state.editingRoomId; });
        if (roomIndex !== -1) {
          state.rooms[roomIndex].name = name;
          state.rooms[roomIndex].capacity = cap;
          state.rooms[roomIndex].priority = priority;
        }
        cancelEditMode();
        showToast('Salon güncellendi.', 'success');
      } else {
        state.rooms.push({
          id: Date.now(),
          name: name,
          capacity: cap,
          priority: priority
        });
        showToast('"' + name + '" salonu eklendi.', 'success');
      }

      state.rooms.sort(function (a, b) { return a.priority - b.priority; });
      nameInput.value = '';
      capInput.value = '';
      priorityInput.value = '';
      updateUI();
      renderRoomTable();
    });

    // Export rooms
    document.getElementById('btn-export-rooms').addEventListener('click', function () {
      if (state.rooms.length === 0) {
        showToast('Dışa aktarılacak salon yok.', 'warning');
        return;
      }
      exportRoomsToExcel(state.rooms);
    });

    // Nav: Start distribution
    document.getElementById('btn-to-step-3').addEventListener('click', function () {
      if (state.rooms.length === 0) {
        showToast('Lütfen en az bir salon tanımlayın.', 'warning');
        return;
      }
      state.rooms.sort(function (a, b) { return a.priority - b.priority; });
      distributeStudents();
      setStep(3);
    });

    // Nav: Step 3 → 2
    document.getElementById('btn-back-to-step-2').addEventListener('click', function () { setStep(2); });

    // Restart
    document.getElementById('btn-restart').addEventListener('click', function () {
      if (confirm('Tüm veriler silinecek ve başa dönülecek. Onaylıyor musunuz?')) {
        location.reload();
      }
    });

    // Export Excel
    document.getElementById('btn-export-excel').addEventListener('click', function () {
      if (!state.distribution) return;
      exportToExcel(state.distribution, state.rooms);
    });

    // Export PDF
    document.getElementById('btn-export-pdf').addEventListener('click', function () {
      if (!state.distribution) return;
      var btn = document.getElementById('btn-export-pdf');
      var originalHTML = btn.innerHTML;
      btn.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2v4m0 12v4M4.93 4.93l2.83 2.83m8.48 8.48l2.83 2.83M2 12h4m12 0h4M4.93 19.07l2.83-2.83m8.48-8.48l2.83-2.83"/></svg> Hazırlanıyor...';
      btn.disabled = true;

      try {
        generatePDF(state.distribution, state.rooms);
      } catch (e) {
        console.error(e);
        showToast('PDF oluşturulurken hata oluştu.', 'error');
      } finally {
        btn.innerHTML = originalHTML;
        btn.disabled = false;
      }
    });

    // Enter key on room inputs
    ['manual-room-name', 'manual-room-capacity', 'manual-room-priority'].forEach(function (id) {
      document.getElementById(id).addEventListener('keydown', function (e) {
        if (e.key === 'Enter') {
          document.getElementById('btn-add-room').click();
        }
      });
    });
  }

  // ===== START =====
  document.addEventListener('DOMContentLoaded', init);

})();
