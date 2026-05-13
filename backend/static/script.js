var failuresData = [];
var fullData = null;
var xolRecords = [];
var currentXolATM = null;
var scriptsCB = [];
var scriptsFallas = [];

// ==========================================
// TOAST & MODAL (Bootstrap)
// ==========================================

function showToast(message, type) {
    type = type || 'success';
    var toast = document.getElementById('appToast');
    var icon = document.getElementById('toast-icon');
    var msg = document.getElementById('toast-message');

    toast.className = 'toast align-items-center border-0 shadow';
    if (type === 'success') {
        toast.classList.add('bg-success-custom');
        icon.className = 'bi bi-check-circle-fill';
    } else if (type === 'danger') {
        toast.classList.add('bg-danger-custom');
        icon.className = 'bi bi-exclamation-triangle-fill';
    } else {
        toast.classList.add('bg-warning');
        icon.className = 'bi bi-info-circle-fill';
    }
    msg.textContent = message;

    var bsToast = new bootstrap.Toast(toast, { delay: 4000 });
    bsToast.show();
}

function showConfirm(message, onConfirm, title) {
    title = title || 'Confirmar';
    document.getElementById('confirmTitle').textContent = title;
    document.getElementById('confirmMessage').textContent = message;
    var modal = new bootstrap.Modal(document.getElementById('confirmModal'));

    var btn = document.getElementById('confirmBtn');
    var newBtn = btn.cloneNode(true);
    btn.parentNode.replaceChild(newBtn, btn);
    newBtn.id = 'confirmBtn';
    newBtn.addEventListener('click', function() {
        modal.hide();
        onConfirm();
    });

    modal.show();
}

// ==========================================
// ESTADO Y CARGA DE DATOS
// ==========================================

function toggleFeriadoIndicator() {
    var checked = document.getElementById('mode-feriado').checked;
    var indicator = document.getElementById('feriado-indicator');
    if (checked) {
        indicator.classList.remove('d-none');
    } else {
        indicator.classList.add('d-none');
    }
}

async function updateStatus() {
    try {
        var res = await fetch('/api/status');
        var data = await res.json();
        var badge = document.getElementById('connection-status');
        if (data.excel_connected) {
            badge.textContent = '\u2713 Planilla Conectada';
            badge.className = 'badge bg-success';
        } else {
            badge.textContent = '\u2717 No Encontrada';
            badge.className = 'badge bg-danger';
        }
        document.getElementById('atm-count').textContent = data.n_atms + ' ATMs';
        
        // Show Sunday indicator automatically (separate from holiday checkbox)
        var domingoIndicator = document.getElementById('domingo-indicator');
        if (domingoIndicator) {
            if (data.es_domingo) {
                domingoIndicator.classList.remove('d-none');
            } else {
                domingoIndicator.classList.add('d-none');
            }
        }
    } catch (e) { console.error(e); }
}

async function loadData() {
    try {
        var res = await fetch('/api/load-data');
        var d = await res.json();
        if (d.status === 'success') {
            fullData = d.data;
            console.log('Datos cargados:', Object.keys(fullData.unificado).length, 'ATMs');
        } else {
            console.error('Error cargando datos:', d.message);
        }
    } catch (e) { console.error('Error fetch load-data:', e); }
}

// ==========================================
// TABS PRINCIPALES
// ==========================================

function showTab(tabId) {
    document.querySelectorAll('.tab-content').forEach(function(t) { t.classList.remove('active'); });
    document.querySelectorAll('#mainTabs .nav-link').forEach(function(b) { b.classList.remove('active'); });
    document.getElementById('tab-' + tabId).classList.add('active');
    event.currentTarget.classList.add('active');

    if (tabId === 'xolusat') {
        loadXolRecords();
    }
}

// ==========================================
// XOLUSAT SUB-TABS
// ==========================================

function showXolTab(subtabId, btn) {
    document.querySelectorAll('.xol-subtab-content').forEach(function(t) { t.classList.remove('active'); });
    document.querySelectorAll('.nav-tabs .nav-link').forEach(function(b) { b.classList.remove('active'); });
    document.getElementById('xol-tab-' + subtabId).classList.add('active');
    btn.classList.add('active');

    if (subtabId === 'lista') {
        loadXolRecords();
    }
}

// ==========================================
// PROCESAR FALLAS
// ==========================================

async function processFailures() {
    var text = document.getElementById('pasted-text').value;
    if (!text) return showToast('Pegá datos primero', 'danger');

    // Asegurar que fullData esté cargado
    if (!fullData) {
        await loadData();
    }

    try {
        var res = await fetch('/api/process-failures', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ text: text })
        });
        var data = await res.json();
        if (data.status === 'success') {
                failuresData = data.failures;
            renderPreview();
            loadScripts();
            showToast(data.failures.length + ' fallas cargadas', 'success');
            var emailInfo = document.getElementById('email-info');
            if (emailInfo) emailInfo.textContent = data.failures.length + ' fallas listas para enviar';
        } else {
            showToast(data.message, 'danger');
        }
    } catch (e) { showToast('Error al procesar datos', 'danger'); }
}

function renderPreview() {
    var tbody = document.querySelector('#preview-table tbody');
    tbody.innerHTML = '';
    var notFound = [];

    failuresData.slice(0, 15).forEach(function(row) {
        var idRaw = row['0'] || '';
        var tipo = row['2'] || '';
        var custodio = row['_custodio'] || '';
        var found = row['_found'] === true || row['_found'] === 'true';

        if (!found && idRaw.trim() && idRaw.toLowerCase() !== 'nan') {
            if (notFound.indexOf(idRaw) === -1) notFound.push(idRaw);
        }

        var custLabel = custodio || '-';
        var custClass = custodioClass(custodio);
        var tr = document.createElement('tr');
        tr.innerHTML = '<td><strong>' + idRaw + '</strong></td><td>' + tipo + '</td><td><span class="badge ' + custClass + '">' + custLabel + '</span></td>';
        tbody.appendChild(tr);
    });

    // Mostrar sección de no encontrados
    var section = document.getElementById('not-found-section');
    var list = document.getElementById('not-found-list');
    if (notFound.length > 0) {
        section.style.display = 'block';
        list.innerHTML = '';
        notFound.forEach(function(id) {
            var div = document.createElement('div');
            div.className = 'border rounded p-3 bg-light';
            div.id = 'notfound-' + normalizarId(id);
            div.innerHTML =
                '<div class="d-flex align-items-center gap-3 mb-2">' +
                    '<span class="badge bg-danger">' + id + '</span>' +
                    '<span class="text-muted small">No registrado en la planilla</span>' +
                '</div>' +
                '<div class="row g-2">' +
                    '<div class="col-md-4">' +
                        '<input type="text" class="form-control form-control-sm" placeholder="Nombre (Ag. Ejemplo)" id="nf-name-' + normalizarId(id) + '">' +
                    '</div>' +
                    '<div class="col-md-4">' +
                        '<select class="form-select form-select-sm" id="nf-cust-' + normalizarId(id) + '">' +
                            '<option value="SUCURSAL">SUCURSAL</option>' +
                            '<option value="STE Metro">STE Metro</option>' +
                            '<option value="Brinks METRO">Brinks METRO</option>' +
                            '<option value="Brinks NORTE">Brinks NORTE</option>' +
                            '<option value="Brinks ESTE">Brinks ESTE</option>' +
                            '<option value="Brinks SUR">Brinks SUR</option>' +
                        '</select>' +
                    '</div>' +
                    '<div class="col-md-2">' +
                        '<input type="text" class="form-control form-control-sm" placeholder="SLA" id="nf-sla-' + normalizarId(id) + '">' +
                    '</div>' +
                    '<div class="col-md-2">' +
                        '<button class="btn btn-success btn-sm w-100" onclick="addNotFoundATM(\'' + id + '\')"><i class="bi bi-plus-circle"></i> Agregar</button>' +
                    '</div>' +
                '</div>';
            list.appendChild(div);
        });
    } else {
        section.style.display = 'none';
    }
}

function normalizarId(id) {
    return (id || '').toUpperCase().replace(/[.\-_\s\/]/g, '');
}

async function addNotFoundATM(idRaw) {
    var idNorm = normalizarId(idRaw);
    var nombre = document.getElementById('nf-name-' + idNorm).value.trim();
    var custodio = document.getElementById('nf-cust-' + idNorm).value;
    var sla = document.getElementById('nf-sla-' + idNorm).value.trim();

    if (!nombre) return showToast('Ingrese el nombre del ATM', 'danger');
    if (!sla) return showToast('Ingrese el SLA', 'danger');

    try {
        var res = await fetch('/api/add-atm', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id: idRaw, nombre: nombre, sla: sla, custodio: custodio })
        });
        var data = await res.json();
        if (data.status === 'success') {
            showToast('ATM ' + idRaw + ' agregado correctamente', 'success');
            // Re-procesar fallas para que el backend devuelva el custodio actualizado
            var text = document.getElementById('pasted-text').value;
            if (text) {
                var res2 = await fetch('/api/process-failures', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ text: text })
                });
                var data2 = await res2.json();
                if (data2.status === 'success') {
                    failuresData = data2.failures;
                }
            }
            await loadData();
            renderPreview();
            loadScripts();
        } else {
            showToast(data.message || 'Error al guardar', 'danger');
        }
    } catch (e) { showToast('Error al guardar ATM', 'danger'); }
}

function custodioClass(custodio) {
    var c = (custodio || '').toUpperCase();
    if (c.indexOf('BRINKS') !== -1) return 'bg-warning text-dark';
    if (c.indexOf('STE') !== -1) return 'bg-info';
    if (c.indexOf('SUCURSAL') !== -1) return 'bg-success';
    return 'bg-secondary';
}

function clearFailures() {
    document.getElementById('pasted-text').value = '';
    failuresData = [];
    scriptsFallas = [];
    renderPreview();
    document.getElementById('scripts-container-fallas').innerHTML = '<div class="text-center text-muted py-5"><i class="bi bi-inbox fs-1 d-block mb-2"></i>Cargá fallas en Gestión de Fallas para generar scripts.</div>';
    var emailInfo = document.getElementById('email-info');
    if (emailInfo) emailInfo.textContent = 'Cargá fallas para comenzar';
    document.getElementById('not-found-section').style.display = 'none';
    showToast('Datos limpiados', 'success');
}

function clearScriptsCB() {
    document.getElementById('cb-search-text').value = '';
    scriptsCB = [];
    renderCbFoundTable([]);
    document.getElementById('scripts-container-cb').innerHTML = '<div class="text-center text-muted py-5"><i class="bi bi-inbox fs-1 d-block mb-2"></i>Generá scripts en Close & Block para verlos aquí.</div>';
    showToast('Scripts C&B limpiados', 'success');
}

// ==========================================
// SCRIPTS
// ==========================================

async function loadScripts() {
    if (failuresData.length === 0) return;
    var isFeriado = document.getElementById('mode-feriado').checked;

    try {
        var res = await fetch('/api/generate-scripts', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ failures: failuresData, is_feriado: isFeriado })
        });
        var data = await res.json();
        if (data.status === 'success') {
            scriptsFallas = data.scripts || [];
            renderScriptsOnlyFallas();
        }
    } catch (e) { console.error(e); }
}

function renderScriptsOnlyFallas() {
    var container = document.getElementById('scripts-container-fallas');
    container.innerHTML = '';
    
    if (scriptsFallas.length === 0) {
        container.innerHTML = '<div class="text-center text-muted py-5">No se generaron scripts de escalamiento.</div>';
        return;
    }
    
    scriptsFallas.forEach(function(s) {
        var div = document.createElement('div');
        div.className = 'script-item';
        var safeComment = s.comentario.replace(/'/g, "\\'").replace(/"/g, '&quot;');
        var destino = extraerDestino(s.comentario);
        var destinoClass = destino === 'Brinks' ? 'bg-warning text-dark' : destino === 'STE' ? 'bg-info' : 'bg-secondary';
        div.innerHTML =
            '<div class="script-header">' +
                '<div><span class="badge bg-dark me-2">TK: ' + s.ticket + '</span><span class="badge ' + destinoClass + '">' + destino + '</span></div>' +
                '<button onclick="copyToClipboard(\'' + safeComment + '\')" class="btn btn-outline-dark btn-sm"><i class="bi bi-clipboard"></i></button>' +
            '</div>' +
            '<code class="small text-dark">' + s.comentario + '</code>';
        container.appendChild(div);
    });
}

function renderScriptsOnlyCB() {
    var container = document.getElementById('scripts-container-cb');
    container.innerHTML = '';
    
    var validScripts = scriptsCB.filter(function(s) {
        var tk = (s.ticket || '').replace(/N\/A/gi, '').trim();
        return tk !== '';
    });
    
    if (validScripts.length === 0) {
        container.innerHTML = '<div class="text-center text-muted py-5">No se generaron scripts de Close & Block.</div>';
        return;
    }
    
    validScripts.forEach(function(s) {
        var div = document.createElement('div');
        div.className = 'script-item';
        var tk = (s.ticket || '').replace(/N\/A/gi, '').trim();
        var cleanComment = (s.comentario || '').replace(/^[\d]+\s*N\/A\s*/, '').trim();
        var safeComment = cleanComment.replace(/'/g, "\\'").replace(/"/g, '&quot;');
        div.innerHTML =
            '<div class="script-header">' +
                '<span class="badge bg-dark">' + tk + '</span>' +
                '<button onclick="copyToClipboard(\'' + safeComment + '\')" class="btn btn-outline-dark btn-sm"><i class="bi bi-clipboard"></i></button>' +
            '</div>' +
            '<code class="small text-dark">' + cleanComment + '</code>';
        container.appendChild(div);
    });
}

function extraerDestino(comentario) {
    var match = comentario.match(/Se escala a (\S+)/);
    return match ? match[1] : 'N/A';
}

async function exportScriptsToExcel() {
    if (failuresData.length === 0) return showToast('No hay datos para exportar', 'danger');
    var isFeriado = document.getElementById('mode-feriado').checked;

    try {
        var resScripts = await fetch('/api/generate-scripts', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ failures: failuresData, is_feriado: isFeriado })
        });
        var dataScripts = await resScripts.json();
        if (dataScripts.status !== 'success') return showToast('Error generando scripts', 'danger');

        var res = await fetch('/api/export-scripts', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ scripts: dataScripts.scripts })
        });

        if (res.ok) {
            var blob = await res.blob();
            var url = window.URL.createObjectURL(blob);
            var a = document.createElement('a');
            a.href = url;
            a.download = 'Scripts_Export_' + new Date().getTime() + '.xlsx';
            document.body.appendChild(a);
            a.click();
            a.remove();
            showToast('Excel exportado', 'success');
        } else {
            showToast('Error al exportar', 'danger');
        }
    } catch (e) { showToast('Error de conexión', 'danger'); }
}

async function exportScriptsCBToExcel() {
    if (scriptsCB.length === 0) return showToast('No hay scripts C&B para exportar', 'danger');

    var validScripts = scriptsCB.filter(function(s) {
        var tk = (s.ticket || '').replace(/N\/A/gi, '').trim();
        return tk !== '';
    });

    if (validScripts.length === 0) return showToast('No hay scripts válidos para exportar', 'danger');

    try {
        var res = await fetch('/api/export-scripts', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ scripts: validScripts })
        });

        if (res.ok) {
            var blob = await res.blob();
            var url = window.URL.createObjectURL(blob);
            var a = document.createElement('a');
            a.href = url;
            a.download = 'Scripts_CB_Export_' + new Date().getTime() + '.xlsx';
            document.body.appendChild(a);
            a.click();
            a.remove();
            showToast('Excel exportado', 'success');
        } else {
            showToast('Error al exportar', 'danger');
        }
    } catch (e) { showToast('Error de conexión', 'danger'); }
}

function copyToClipboard(text) {
    navigator.clipboard.writeText(text).then(function() {
        showToast('Copiado al portapapeles', 'success');
    });
}

// ==========================================
// ENVIAR CORREOS
// ==========================================

async function sendEmails() {
    if (failuresData.length === 0) return showToast('Cargá fallas primero', 'danger');
    var isFeriado = document.getElementById('mode-feriado').checked;

    try {
        var res = await fetch('/api/send-emails', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ failures: failuresData, is_feriado: isFeriado })
        });
        var data = await res.json();
        if (data.status === 'success') {
            var r = data.results;
            showToast('Correos: ' + r.abiertos + ' abiertos | ' + r.sin_sla + ' sin SLA | ' + r.sin_contacto + ' sin contacto', 'success');
        } else {
            showToast(data.message || 'Error al enviar correos', 'danger');
        }
    } catch (e) { showToast('Error al enviar correos', 'danger'); }
}

// ==========================================
// XOLUSAT
// ==========================================

async function searchATMXol() {
    var id = document.getElementById('xol-search-id').value;
    var infoContainer = document.getElementById('xol-atm-info');
    if (!id) return;

    try {
        var res = await fetch('/api/xolusat/search', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id: id })
        });
        var data = await res.json();

        if (data.status === 'found') {
            currentXolATM = data;
            infoContainer.innerHTML = '<div class="alert alert-success py-2 mb-0 d-flex align-items-center gap-2"><i class="bi bi-check-circle-fill"></i> <strong>' + data.nombre + '</strong> <span class="badge bg-success">' + data.sla + '</span></div>';
            document.getElementById('xol-sla').value = data.sla || '';
            document.getElementById('xol-atm-nombre').value = data.nombre || '';
        } else {
            currentXolATM = null;
            infoContainer.innerHTML = '<div class="alert alert-danger py-2 mb-0 d-flex align-items-center gap-2"><i class="bi bi-exclamation-triangle-fill"></i> ATM no encontrado</div>';
            document.getElementById('xol-sla').value = '';
            document.getElementById('xol-atm-nombre').value = '';
        }
    } catch (e) { console.error(e); }
}

function clearXolSearch() {
    document.getElementById('xol-search-id').value = '';
    document.getElementById('xol-atm-info').innerHTML = '';
    document.getElementById('xol-sla').value = '';
    document.getElementById('xol-atm-nombre').value = '';
    document.getElementById('xol-incident').value = '';
    document.getElementById('xol-detalle').value = '';
    currentXolATM = null;
}

async function registerXol(sendEmail) {
    var incident = document.getElementById('xol-incident').value;
    var estado = document.getElementById('xol-estado').value;
    var subcatSelect = document.getElementById('xol-subcat');
    var subcat = subcatSelect.value;
    if (subcat === 'OTRA') {
        subcat = document.getElementById('xol-subcat-otra').value;
        if (!subcat.trim()) return showToast('Especifique la subcategoría', 'danger');
    }
    var detalle = document.getElementById('xol-detalle').value;
    var atmId = document.getElementById('xol-search-id').value;
    var sla = document.getElementById('xol-sla').value;
    var atmNombre = document.getElementById('xol-atm-nombre').value;

    if (!incident || !atmId) return showToast('Incident e ID ATM son obligatorios', 'danger');

    var custodio = currentXolATM ? currentXolATM.custodio : 'SUCURSAL';

    try {
        var res = await fetch('/api/xolusat/register', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                incident: incident, estado: estado, id_atm: atmId,
                subcategoria: subcat, detalle: detalle, sla: sla,
                atm_nombre: atmNombre, custodio: custodio, send_email: sendEmail
            })
        });
        var data = await res.json();
        if (data.status === 'success' || data.status === 'warning') {
            showToast(data.message, 'success');
            loadXolRecords();
        } else {
            showToast(data.message, 'danger');
        }
    } catch (e) { showToast('Error al registrar', 'danger'); }
}

async function loadXolRecords() {
    var estadoFilter = document.getElementById('xol-filter-estado');
    var subcatFilter = document.getElementById('xol-filter-subcat');
    var params = '';
    if (estadoFilter && estadoFilter.value !== 'Todos') {
        params += '?estado=' + encodeURIComponent(estadoFilter.value);
    }
    if (subcatFilter && subcatFilter.value !== 'Todas') {
        params += (params ? '&' : '?') + 'subcategoria=' + encodeURIComponent(subcatFilter.value);
    }

    try {
        var res = await fetch('/api/xolusat/list' + params);
        var data = await res.json();
        if (data.status === 'success') {
            xolRecords = data.records;
            renderXolTable();
            updateIncidentDropdown();
        }
    } catch (e) { console.error(e); }
}

function renderXolTable() {
    var tbody = document.querySelector('#xol-table tbody');
    tbody.innerHTML = '';
    xolRecords.forEach(function(r) {
        var tr = document.createElement('tr');
        tr.innerHTML =
            '<td><strong>' + (r.incident || '') + '</strong></td>' +
            '<td>' + estadoBadge(r.estado) + '</td>' +
            '<td>' + (r.id_atm || '') + '</td>' +
            '<td>' + (r.subcategoria || '') + '</td>' +
            '<td class="text-truncate" style="max-width:120px;">' + (r.detalle || '') + '</td>' +
            '<td>' + (r.sla || '') + '</td>' +
            '<td>' + (r.atm_nombre || '') + '</td>' +
            '<td class="text-truncate" style="max-width:100px;">' + (r.custodio || '') + '</td>' +
            '<td class="text-muted small">' + (r.fecha_reg || '') + '</td>';
        tbody.appendChild(tr);
    });
}

function estadoBadge(estado) {
    var map = { 'OPEN': 'bg-success', 'SUPEND': 'bg-warning text-dark', 'DISPACHED': 'bg-info', 'CLOSED': 'bg-secondary' };
    var cls = map[estado] || 'bg-secondary';
    return '<span class="badge ' + cls + '">' + estado + '</span>';
}

function updateIncidentDropdown() {
    var select = document.getElementById('xol-update-incident');
    select.innerHTML = '';
    xolRecords.forEach(function(r) {
        var opt = document.createElement('option');
        opt.value = r.incident;
        opt.text = r.incident + ' - ' + r.id_atm;
        select.appendChild(opt);
    });
}

function toggleSubcatOtra() {
    var select = document.getElementById('xol-subcat');
    var otraInput = document.getElementById('xol-subcat-otra-container');
    otraInput.style.display = select.value === 'OTRA' ? 'block' : 'none';
}

async function updateXolStatus() {
    var incident = document.getElementById('xol-update-incident').value;
    var estado = document.getElementById('xol-update-estado').value;
    if (!incident) return showToast('Seleccione un incidente', 'danger');

    try {
        var res = await fetch('/api/xolusat/update-status', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ incident: incident, estado: estado })
        });
        var data = await res.json();
        showToast(data.message, 'success');
        loadXolRecords();
    } catch (e) { showToast('Error al actualizar', 'danger'); }
}

function clearXolRecords() {
    showConfirm('¿Eliminar todos los registros XOLUSAT?', async function() {
        try {
            await fetch('/api/xolusat/clear', { method: 'POST' });
            xolRecords = [];
            renderXolTable();
            updateIncidentDropdown();
            showToast('Registros limpiados', 'success');
        } catch (e) { console.error(e); }
    }, 'Limpiar registros');
}

// ==========================================
// RCU UPLOAD
// ==========================================

document.getElementById('rcu-upload').addEventListener('change', async function(e) {
    var file = e.target.files[0];
    if (!file) return;

    var statusDiv = document.getElementById('rcu-status');
    var btn = document.getElementById('rcu-btn');

    // Estado: cargando
    statusDiv.style.display = 'block';
    statusDiv.innerHTML =
        '<div class="alert alert-info py-2 mb-0">' +
            '<div class="d-flex align-items-center gap-2">' +
                '<div class="spinner-border spinner-border-sm" role="status"></div>' +
                '<div>' +
                    '<strong>' + file.name + '</strong><br>' +
                    '<small>Procesando...</small>' +
                '</div>' +
            '</div>' +
        '</div>';
    btn.disabled = true;
    btn.innerHTML = '<div class="spinner-border spinner-border-sm me-1"></div> Procesando...';

    var formData = new FormData();
    formData.append('file', file);

    try {
        var res = await fetch('/api/upload-rcu', { method: 'POST', body: formData });
        var data = await res.json();

        if (data.status === 'success' && data.results) {
            var r = data.results;
            statusDiv.innerHTML =
                '<div class="alert alert-success py-2 mb-0">' +
                    '<i class="bi bi-check-circle-fill me-1"></i> <strong>Completado</strong>' +
                    '<div class="mt-1 small">' +
                        '<div><i class="bi bi-arrow-repeat me-1"></i> Actualizados: <strong>' + r.actualizados + '</strong></div>' +
                        '<div><i class="bi bi-plus-circle me-1"></i> Nuevos: <strong>' + r.nuevos + '</strong></div>' +
                        '<div class="text-muted"><i class="bi bi-files me-1"></i> Total procesados: ' + r.total_procesados + ' de ' + r.total_archivo + '</div>' +
                    '</div>' +
                '</div>';
            showToast('RCU procesado: ' + r.actualizados + ' actualizados, ' + r.nuevos + ' nuevos', 'success');
        } else if (data.status === 'success') {
            statusDiv.innerHTML =
                '<div class="alert alert-success py-2 mb-0">' +
                    '<i class="bi bi-check-circle-fill me-1"></i> ' + (data.message || 'Archivo procesado') +
                '</div>';
            showToast(data.message || 'RCU procesado', 'success');
        } else {
            statusDiv.innerHTML =
                '<div class="alert alert-danger py-2 mb-0">' +
                    '<i class="bi bi-exclamation-triangle-fill me-1"></i> ' + (data.message || 'Error al procesar') +
                '</div>';
            showToast(data.message || 'Error al procesar RCU', 'danger');
        }
        updateStatus();
        loadData();
    } catch (err) {
        statusDiv.innerHTML =
            '<div class="alert alert-danger py-2 mb-0">' +
                '<i class="bi bi-exclamation-triangle-fill me-1"></i> Error de conexión' +
            '</div>';
        showToast('Error al subir archivo', 'danger');
    } finally {
        btn.disabled = false;
        btn.innerHTML = '<i class="bi bi-upload me-1"></i> Seleccionar archivo';
        // Limpiar el input para permitir re-subir el mismo archivo
        e.target.value = '';
    }
});

// ==========================================
// CLOSED AND BLOCK
// ==========================================

// Normalizar IDs de ATMs en texto (RD05327 -> BHD05327)
function normalizarATMsEnTexto(texto) {
    if (!texto) return texto;
    // Reemplazar RD0 por BHD (RD05332 -> BHD05332)
    return texto.replace(/RD0/gi, 'BHD');
}

async function cbAgregar() {
    var textRaw = document.getElementById('cb-text').value.trim();
    var asunto = document.getElementById('cb-asunto').value.trim();
    var reportadoPor = document.getElementById('cb-reportado-por').value.trim();

    if (!textRaw) return showToast('Ingresá IDs primero', 'danger');

    // Normalizar IDs de ATMs (RD05327 -> BHD05327)
    var text = normalizarATMsEnTexto(textRaw);

    try {
        var res = await fetch('/api/closed-block/agregar', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ text: text, asunto: asunto, reportado_por: reportadoPor })
        });
        var data = await res.json();
        if (data.status === 'success') {
            showToast(data.message, 'success');
            document.getElementById('cb-text').value = '';
            document.getElementById('cb-asunto').value = '';
            document.getElementById('cb-reportado-por').value = '';
        } else {
            showToast(data.message || 'Error', 'danger');
        }
    } catch (e) { showToast('Error de conexión', 'danger'); }
}

function cbLimpiarInput() {
    document.getElementById('cb-text').value = '';
    document.getElementById('cb-asunto').value = '';
    document.getElementById('cb-reportado-por').value = '';
    showToast('Área limpiada', 'success');
}

async function cbBuscarCoincidencias() {
    var text = document.getElementById('cb-search-text').value.trim();
    if (!text) return showToast('Pegá datos de fallas primero', 'danger');

    try {
        var res = await fetch('/api/closed-block/buscar', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ text: text })
        });
        var data = await res.json();
        if (data.status === 'success') {
            renderCbFoundTable(data.found || []);
        } else {
            showToast(data.message || 'Error', 'danger');
        }
    } catch (e) { showToast('Error de conexión', 'danger'); }
}

function renderCbFoundTable(found) {
    var tbody = document.getElementById('cb-found-tbody');
    var emptyDiv = document.getElementById('cb-found-empty');
    var tableContainer = document.getElementById('cb-found-table').parentElement;
    var headerDiv = document.getElementById('cb-found-header');
    var countSpan = document.getElementById('cb-found-count');

    tbody.innerHTML = '';

    if (!found || found.length === 0) {
        emptyDiv.style.display = 'block';
        tableContainer.style.display = 'none';
        headerDiv.style.display = 'none';
        return;
    }

    emptyDiv.style.display = 'none';
    tableContainer.style.display = 'block';
    headerDiv.style.display = 'block';
    countSpan.textContent = found.length;

    found.forEach(function(r) {
        var horas = parseInt(r.horas) || 0;
        var colorClass = horas >= 48 ? 'bg-danger' : horas >= 36 ? 'bg-warning text-dark' : 'bg-secondary';
        var tr = document.createElement('tr');
        tr.className = 'table-danger';
        tr.innerHTML =
            '<td><strong>' + r.id + '</strong></td>' +
            '<td class="small text-truncate" style="max-width:120px;">' + (r.nombre || '-') + '</td>' +
            '<td class="small">' + (r.custodio || '-') + '</td>' +
            '<td><span class="badge ' + colorClass + '">' + r.horas + '</span></td>' +
            '<td class="small text-muted">' + (r.fecha || '-') + '</td>' +
            '<td class="small text-truncate" style="max-width:100px;" title="' + (r.asunto || '') + '">' + (r.asunto || '-') + '</td>' +
            '<td class="small">' + (r.reportado_por || '-') + '</td>';
        tbody.appendChild(tr);
    });
}

function cbLimpiarBusqueda() {
    document.getElementById('cb-search-text').value = '';
    renderCbFoundTable([]);
    showToast('Búsqueda limpiada', 'success');
}

async function cbGenerarScripts() {
    var text = document.getElementById('cb-search-text').value.trim();
    if (!text) return showToast('Pegá datos primero', 'danger');

    try {
        var res = await fetch('/api/generate-scripts-cb', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ text: text })
        });
        var data = await res.json();
        if (data.status === 'success') {
            scriptsCB = data.scripts || [];
            renderScriptsOnlyCB();
            showToast(scriptsCB.length + ' scripts C&B generados', 'success');
            showTab('scripts');
        } else {
            showToast(data.message || 'Error', 'danger');
        }
    } catch (e) { showToast('Error de conexión', 'danger'); }
}

function toggleCollapse(id) {
    var el = document.getElementById(id);
    var icon = document.getElementById(id + '-icon');
    if (el.style.display === 'none') {
        el.style.display = 'block';
        if (icon) icon.className = 'bi bi-chevron-down';
    } else {
        el.style.display = 'none';
        if (icon) icon.className = 'bi bi-chevron-right';
    }
}

// ==========================================
// INIT
// ==========================================

updateStatus();
loadData();
setInterval(updateStatus, 30000);

