let currentRole = '';
let repartoData = [];
let adminTab = 'create';

document.addEventListener("DOMContentLoaded", function() {
    const dateField = document.getElementById('hdrFecha');
    if(dateField) dateField.value = new Date().toISOString().slice(0, 10);
});


// Muestra el cuadro de contraseña cuando el usuario hace clic en Administrador
function requestAdminAccess() {
    const passwordBox = document.getElementById('adminPasswordBox');
    passwordBox.classList.toggle('hidden');
    if(!passwordBox.classList.contains('hidden')) {
        document.getElementById('adminPasswordInput').focus();
    }
}

// Valida la clave por defecto "admin"
function verifyAdminPassword() {
    const pass = document.getElementById('adminPasswordInput').value.trim();
    if (pass === "admin") {
        selectRole('admin');
    } else {
        alert("Contraseña incorrecta.");
        document.getElementById('adminPasswordInput').value = '';
    }
}

function selectRole(role) {
    currentRole = role;
    document.getElementById('roleSelectorScreen').classList.add('hidden');
    document.getElementById('badgeRol').textContent = role === 'admin' ? 'Administración' : 'Chofer';

    if (role === 'admin') {
        document.getElementById('adminMenu').classList.remove('hidden');
        switchAdminTab('create');
    } else {
        document.getElementById('generalSearchContainer').classList.remove('hidden');
        document.getElementById('driverModal').classList.remove('hidden');
    }
}

async function switchAdminTab(tab) {
    adminTab = tab;
    const btnCreate = document.getElementById('btnTabCreate');
    const btnMonitor = document.getElementById('btnTabMonitor');
    const createControls = document.getElementById('adminCreateControls');
    const monitorBox = document.getElementById('monitorSelectorBox');
    const searchContainer = document.getElementById('generalSearchContainer');

    const makeEditable = (tab === 'create');
    ['hdrNumero', 'hdrChofer', 'hdrZona', 'hdrPatente'].forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            if (makeEditable) {
                el.removeAttribute('readonly');
                el.classList.remove('bg-gray-50');
                el.classList.add('bg-white');
            } else {
                el.setAttribute('readonly', true);
                el.classList.add('bg-gray-50');
                el.classList.remove('bg-white');
            }
        }
    });

    if (tab === 'create') {
        btnCreate.className = "bg-blue-600 text-white px-3 py-2 rounded-lg text-xs font-bold flex items-center gap-2 transition";
        btnMonitor.className = "bg-gray-200 text-gray-700 px-3 py-2 rounded-lg text-xs font-bold flex items-center gap-2 transition";
        createControls.classList.remove('hidden');
        monitorBox.classList.add('hidden');
        searchContainer.classList.add('hidden');
        repartoData = [];
        renderAll();
    } else {
        btnMonitor.className = "bg-blue-600 text-white px-3 py-2 rounded-lg text-xs font-bold flex items-center gap-2 transition";
        btnCreate.className = "bg-gray-200 text-gray-700 px-3 py-2 rounded-lg text-xs font-bold flex items-center gap-2 transition";
        createControls.classList.add('hidden');
        monitorBox.classList.remove('hidden');
        searchContainer.classList.remove('hidden');
    }
}
// Función para que el administrador exporte la ruta que está monitoreando en Excel
function exportMonitorExcel() {
    if (repartoData.length === 0) return alert("No hay datos de ruta cargados para exportar.");
    
    const hdrNum = document.getElementById('hdrNumero').value || 'REPORTE';
    const chofer = document.getElementById('hdrChofer').value || 'S_D';

    // Convertir los datos actuales al formato de exportación ordenado
    const tableRows = repartoData.map(i => ({
        "HDR": i.HDR,
        "CHOFER": i.CHOFER,
        "ZONA": i.ZONA,
        "PATENTE": i.PATENTE,
        "COD_CLIENTE": i.COD_CLIENTE,
        "RAZON_SOCIAL": i.RAZON_SOCIAL,
        "DIRECCION": i.DIRECCION,
        "VEN": i.VEN,
        "COMPROBANTE": i.COMPROBANTE,
        "IMPORTE": i.IMPORTE,
        "ESTADO_ENTREGA": i.ESTADO_ENTREGA,
        "FORMA_PAGO": i.FORMA_PAGO,
        "OBSERVACIONES": i.OBSERVACIONES,
        "FECHA": i.FECHA || new Date().toISOString().slice(0, 10)
    }));

    const ws = XLSX.utils.json_to_sheet(tableRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Monitoreo_HDR");
    XLSX.writeFile(wb, `Rendicion_HDR_${hdrNum}_${chofer}.xlsx`);
}

// Asegurar que al cargar el archivo de Excel del Administrador se asigne del momento
document.getElementById('fileInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(evt) {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws, { defval: '' });

        const hdrInput = document.getElementById('hdrNumero').value.trim() || 'S/N';
        const choferInput = document.getElementById('hdrChofer').value.trim() || 'S/D';
        const zonaInput = document.getElementById('hdrZona').value.trim() || 'S/D';
        const patenteInput = document.getElementById('hdrPatente').value.trim() || 'S/D';
        // Genera la fecha actual en formato DD/MM/YYYY
        const now = new Date();
        const day = String(now.getDate()).padStart(2, '0');
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const year = now.getFullYear();
        const fechaHoy = `${day}/${month}/${year}`;


        repartoData = json.map(r => {
            let rawCliente = String(r.CLIENTE || r.COD || r.ID || '').trim();
            let clienteFormatted = rawCliente ? rawCliente.padStart(10, '0') : '0000000000';
            let imp = cleanNumber(r.IMPORTE || r.IMPORTA || r.MONTO || 0);

            return {
                HDR: hdrInput,
                CHOFER: choferInput,
                ZONA: zonaInput,
                PATENTE: patenteInput,
                COD_CLIENTE: clienteFormatted,
                RAZON_SOCIAL: String(r['RAZON SOCIAL'] || r.NOMBRE || r.RS || '').trim(),
                DIRECCION: String(r.DIRECCION || '').trim(),
                VEN: String(r.VEN || '').trim(),
                COMPROBANTE: String(r.COMPROBANTE || '').trim(),
                IMPORTE: imp,
                ESTADO_ENTREGA: 'PENDIENTE',
                FORMA_PAGO: 'EFECTIVO',
                OBSERVACIONES: '',
                FECHA: fechaHoy
            };
        });

        renderAll();
        alert("Excel cargado correctamente con fecha de hoy. Presione 'Registrar HDR en Nube'.");
    };
    reader.readAsArrayBuffer(file);
});

// Actualiza también el mapeo mapCloudData para capturar la fecha desde la base de datos
function mapCloudData(rawData) {
    return rawData.map(item => {
        return {
            rowIndex: item.rowIndex,
            HDR: item.HDR || '',
            CHOFER: item.CHOFER || '',
            ZONA: item.ZONA || '',
            PATENTE: item.PATENTE || '',
            COD_CLIENTE: item.COD_CLIENTE || item.CLIENTE || item.COD || '',
            RAZON_SOCIAL: item.RAZON_SOCIAL || item['RAZON SOCIAL'] || '',
            DIRECCION: item.DIRECCION || '',
            VEN: item.VEN || '',
            COMPROBANTE: item.COMPROBANTE || '',
            IMPORTE: cleanNumber(item.IMPORTE || 0),
            ESTADO_ENTREGA: item.ESTADO_ENTREGA || item['ESTADO ENTREGA'] || item.ESTADO || 'PENDIENTE',
            FORMA_PAGO: item.FORMA_PAGO || item['FORMA DE PAGO'] || 'EFECTIVO',
            OBSERVACIONES: item.OBSERVACIONES || '',
            FECHA: item.FECHA || ''
        };
    });
}

async function searchMonitorHDR() {
    const hdr = document.getElementById('inputMonitorHDR').value.trim();
    if (!hdr) return alert("Ingrese un número de HDR para consultar.");
    
    try {
        const res = await fetch(`${SCRIPT_URL}?action=getHDR&hdr=${hdr}`);
        const json = await res.json();
        if(json.status === 'success') {
            repartoData = mapCloudData(json.data);
            if(repartoData.length > 0) {
                document.getElementById('hdrNumero').value = repartoData[0].HDR || hdr;
                document.getElementById('hdrChofer').value = repartoData[0].CHOFER || '';
                document.getElementById('hdrZona').value = repartoData[0].ZONA || '';
                document.getElementById('hdrPatente').value = repartoData[0].PATENTE || '';
            } else {
                alert("No se encontraron registros para la HDR " + hdr);
            }
            renderAll();
        }
    } catch(e) {
        alert("Error al consultar ruta: " + e);
    }
}

function loadDriverHDR() {
    const hdr = document.getElementById('inputHdrLogin').value.trim();
    if (!hdr) return alert("Ingrese un número de HDR válido.");
    document.getElementById('driverModal').classList.add('hidden');
    document.getElementById('hdrNumero').value = hdr;
    loadHDRDataCloud(hdr);
}

async function loadHDRDataCloud(hdr) {
    try {
        const res = await fetch(`${SCRIPT_URL}?action=getHDR&hdr=${hdr}`);
        const json = await res.json();
        if(json.status === 'success') {
            repartoData = mapCloudData(json.data);
            if(repartoData.length > 0) {
                document.getElementById('hdrChofer').value = repartoData[0].CHOFER || '';
                document.getElementById('hdrZona').value = repartoData[0].ZONA || '';
                document.getElementById('hdrPatente').value = repartoData[0].PATENTE || '';
            } else {
                alert("No se encontraron registros para la HDR " + hdr);
            }
            renderAll();
        }
    } catch(e) {
        alert("Error al cargar la ruta: " + e);
    }
}

async function updateCloudRow(index) {
    const item = repartoData[index];
    try {
        await fetch(SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify({
                action: 'updateRow',
                rowIndex: item.rowIndex,
                estado: item.ESTADO_ENTREGA,
                formaPago: item.FORMA_PAGO,
                observacion: item.OBSERVACIONES
            })
        });
    } catch(e) {
        console.error("Error al sincronizar en tiempo real", e);
    }
}

function cleanNumber(val) {
    if (typeof val === 'number') return val;
    if (!val) return 0;
    let str = String(val).replace(/\$/g, '').trim();
    if (str.includes('.') && str.includes(',')) str = str.replace(/\./g, '').replace(',', '.');
    else if (str.includes(',')) str = str.replace(',', '.');
    return parseFloat(str) || 0;
}

function formatMoney(num) {
    return '$ ' + num.toLocaleString('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

document.getElementById('fileInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(evt) {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws, { defval: '' });

        const hdrInput = document.getElementById('hdrNumero').value.trim() || 'S/N';
        const choferInput = document.getElementById('hdrChofer').value.trim() || 'S/D';
        const zonaInput = document.getElementById('hdrZona').value.trim() || 'S/D';
        const patenteInput = document.getElementById('hdrPatente').value.trim() || 'S/D';

        repartoData = json.map(r => {
            let rawCliente = String(r.CLIENTE || r.COD || r.ID || '').trim();
            let clienteFormatted = rawCliente ? rawCliente.padStart(10, '0') : '0000000000';
            let imp = cleanNumber(r.IMPORTE || r.IMPORTA || r.MONTO || 0);

            return {
                HDR: hdrInput,
                CHOFER: choferInput,
                ZONA: zonaInput,
                PATENTE: patenteInput,
                COD_CLIENTE: clienteFormatted,
                RAZON_SOCIAL: String(r['RAZON SOCIAL'] || r.NOMBRE || r.RS || '').trim(),
                DIRECCION: String(r.DIRECCION || '').trim(),
                VEN: String(r.VEN || '').trim(),
                COMPROBANTE: String(r.COMPROBANTE || '').trim(),
                IMPORTE: imp,
                ESTADO_ENTREGA: 'PENDIENTE',
                FORMA_PAGO: 'EFECTIVO',
                OBSERVACIONES: ''
            };
        });

        renderAll();
        alert("Excel cargado correctamente. Verifique los datos y presione 'Registrar HDR en Nube'.");
    };
    reader.readAsArrayBuffer(file);
});

async function registerHDRCloud() {
    if (repartoData.length === 0) return alert("No hay datos cargados para registrar.");
    
    const hdrVal = document.getElementById('hdrNumero').value.trim();
    const choferVal = document.getElementById('hdrChofer').value.trim();
    const zonaVal = document.getElementById('hdrZona').value.trim();
    const patenteVal = document.getElementById('hdrPatente').value.trim();

    if (!hdrVal) return alert("Por favor, ingrese el número de HDR antes de registrar.");

    repartoData.forEach(i => {
        i.HDR = hdrVal;
        i.CHOFER = choferVal;
        i.ZONA = zonaVal;
        i.PATENTE = patenteVal;
    });

    try {
        const res = await fetch(SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'saveAll', rows: repartoData })
        });
        const json = await res.json();
        if(json.status === 'success') {
            alert("¡Hoja de Ruta " + hdrVal + " registrada en Google Sheets con éxito!");
            switchAdminTab('monitor');
        } else {
            alert("Error al registrar: " + json.message);
        }
    } catch(e) {
        alert("Error de conexión al registrar: " + e);
    }
}

function renderAll(filteredData = null) {
    const list = filteredData || repartoData;
    renderMobileCards(list);
    renderDesktopTable(list);
    calculateTotals();
}

// Colores y bordes para los 6 estados
function getStatusBorderClass(estado) {
    if (estado === 'ENTREGADO') return 'border-l-4 border-l-green-600';
    if (estado === 'PENDIENTE_TRANSF') return 'border-l-4 border-l-sky-500';
    if (estado === 'RECHAZO_PARCIAL') return 'border-l-4 border-l-neutral-500';
    if (estado === 'RECHAZO_TOTAL') return 'border-l-4 border-l-red-600';
    if (estado === 'REASIGNADO') return 'border-l-4 border-l-fuchsia-600';
    return 'border-l-4 border-l-amber-400'; // PENDIENTE
}

function getStatusSelectClass(estado) {
    if (estado === 'ENTREGADO') return 'bg-green-50 text-green-700 border-green-300';
    if (estado === 'PENDIENTE_TRANSF') return 'bg-sky-50 text-sky-700 border-sky-300';
    if (estado === 'RECHAZO_PARCIAL') return 'bg-neutral-50 text-neutal-500 border-neutral-300';
    if (estado === 'RECHAZO_TOTAL') return 'bg-red-50 text-red-700 border-red-300';
    if (estado === 'REASIGNADO') return 'bg-fuchsia-50 text-fuchsia-700 border-fuchsia-300';
    return 'bg-amber-50 text-amber-700 border-amber-300'; // PENDIENTE
}

function renderMobileCards(list) {
    const container = document.getElementById('cardsContainer');
    container.innerHTML = '';

    if (list.length === 0) {
        container.innerHTML = '<div class="text-center py-8 bg-white rounded-xl border text-gray-400 text-xs">No hay registros para mostrar.</div>';
        return;
    }

    list.forEach((item, index) => {
        const card = document.createElement('div');
        card.className = `bg-white p-3 rounded-xl border border-gray-200 shadow-xs flex flex-col gap-2 ${getStatusBorderClass(item.ESTADO_ENTREGA)}`;

        card.innerHTML = `
            <div class="flex justify-between items-start">
                <span class="bg-gray-100 text-gray-700 px-2 py-0.5 rounded text-[10px] font-bold font-mono">ID: ${item.COD_CLIENTE}</span>
                <span class="text-sm font-extrabold text-slate-900">${formatMoney(item.IMPORTE)}</span>
            </div>
            <div>
                <div class="text-xs font-bold text-slate-900">${item.RAZON_SOCIAL}</div>
                <div class="text-[11px] text-gray-500"><i class="fa-solid fa-location-dot text-gray-400 mr-1"></i>${item.DIRECCION}</div>
            </div>
            <div class="flex items-center gap-2 pt-1 border-t border-gray-100">
                <span class="bg-gray-50 border px-2 py-0.5 rounded text-[10px] font-semibold text-gray-600"><i class="fa-solid fa-file-lines mr-1"></i>${item.COMPROBANTE}</span>
                <span class="text-[10px] font-semibold text-gray-500">Ven: ${item.VEN || '-'}</span>
            </div>
            <div class="grid grid-cols-2 gap-2 pt-1">
                <select class="..." onchange="updateStatus(${index}, this.value)">
                    <option value="PENDIENTE" ${item.ESTADO_ENTREGA === 'PENDIENTE' ? 'selected' : ''}>🟡 PENDIENTE</option>
                    <option value="ENTREGADO" ${item.ESTADO_ENTREGA === 'ENTREGADO' ? 'selected' : ''}>🟢 ENTREGADO</option>
                    <option value="PENDIENTE_TRANSF" ${item.ESTADO_ENTREGA === 'PENDIENTE_TRANSF' ? 'selected' : ''}>🔵 TRANSF. PEND.</option>
                    <option value="RECHAZO_PARCIAL" ${item.ESTADO_ENTREGA === 'RECHAZO_PARCIAL' ? 'selected' : ''}>⚪ RECHAZO PARCIAL</option>
                    <option value="RECHAZO_TOTAL" ${item.ESTADO_ENTREGA === 'RECHAZO_TOTAL' ? 'selected' : ''}>🔴 RECHAZO TOTAL</option>
                    <option value="REASIGNADO" ${item.ESTADO_ENTREGA === 'REASIGNADO' ? 'selected' : ''}>🟣 REASIGNADO</option>
                </select>
                <select class="w-full p-1.5 rounded-lg border border-gray-300 text-xs font-bold outline-none bg-white cursor-pointer" onchange="updatePayment(${index}, this.value)">
                    <option value="EFECTIVO" ${item.FORMA_PAGO === 'EFECTIVO' ? 'selected' : ''}>💵 EFECTIVO</option>
                    <option value="TRANSFERENCIA" ${item.FORMA_PAGO === 'TRANSFERENCIA' ? 'selected' : ''}>📲 TRANSF.</option>
                    <option value="CHEQUE" ${item.FORMA_PAGO === 'CHEQUE' ? 'selected' : ''}>📜 CHEQUE</option>
                    <option value="CUENTA CORRIENTE" ${item.FORMA_PAGO === 'CUENTA CORRIENTE' ? 'selected' : ''}>📋 CTA CTE</option>
                </select>
            </div>
            <div>
                <input type="text" class="w-full bg-gray-50 border rounded-lg px-2.5 py-1.5 text-xs focus:outline-none" value="${item.OBSERVACIONES || ''}" 
                       onchange="updateObs(${index}, this.value)" placeholder="Añadir nota / novedad...">
            </div>
        `;
        container.appendChild(card);
    });
}

function renderDesktopTable(list) {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';

    list.forEach((item, index) => {
        const tr = document.createElement('tr');
        tr.className = "hover:bg-gray-50 transition";
        tr.innerHTML = `
            <td class="p-2.5 font-bold font-mono text-gray-700">${item.COD_CLIENTE}</td>
            <td class="p-2.5 font-medium text-gray-900">${item.RAZON_SOCIAL}</td>
            <td class="p-2.5 text-gray-600">${item.DIRECCION}</td>
            <td class="p-2.5 text-gray-600">${item.VEN || '-'}</td>
            <td class="p-2.5 text-gray-600">${item.COMPROBANTE}</td>
            <td class="p-2.5 font-bold text-slate-900">${formatMoney(item.IMPORTE)}</td>
            <td class="p-2.5">
                <select class="p-1 rounded border text-xs font-bold outline-none cursor-pointer ${getStatusSelectClass(item.ESTADO_ENTREGA)}" onchange="updateStatus(${index}, this.value)">
                    <option value="PENDIENTE" ${item.ESTADO_ENTREGA === 'PENDIENTE' ? 'selected' : ''}>🟡 PENDIENTE</option>
                    <option value="ENTREGADO" ${item.ESTADO_ENTREGA === 'ENTREGADO' ? 'selected' : ''}>🟢 ENTREGADO</option>
                    <option value="PENDIENTE_TRANSF" ${item.ESTADO_ENTREGA === 'PENDIENTE_TRANSF' ? 'selected' : ''}>🔵 TRANSF. PEND.</option>
                    <option value="RECHAZO_PARCIAL" ${item.ESTADO_ENTREGA === 'RECHAZO_PARCIAL' ? 'selected' : ''}>⚪ RECHAZO PARCIAL</option>
                    <option value="RECHAZO_TOTAL" ${item.ESTADO_ENTREGA === 'RECHAZO_TOTAL' ? 'selected' : ''}>🔴 RECHAZO TOTAL</option>
                    <option value="REASIGNADO" ${item.ESTADO_ENTREGA === 'REASIGNADO' ? 'selected' : ''}>🟣 REASIGNADO</option>
                </select>
            </td>
            <td class="p-2.5">
                <select class="p-1 rounded border border-gray-300 text-xs font-bold outline-none bg-white cursor-pointer" onchange="updatePayment(${index}, this.value)">
                     <option value="EFECTIVO" ${item.FORMA_PAGO === 'EFECTIVO' ? 'selected' : ''}>💵 EFECTIVO</option>
                    <option value="TRANSFERENCIA" ${item.FORMA_PAGO === 'TRANSFERENCIA' ? 'selected' : ''}>📲 TRANSF.</option>
                    <option value="CHEQUE" ${item.FORMA_PAGO === 'CHEQUE' ? 'selected' : ''}>📜 CHEQUE</option>
                    <option value="CUENTA CORRIENTE" ${item.FORMA_PAGO === 'CUENTA CORRIENTE' ? 'selected' : ''}>📋 CTA CTE</option
                </select>
            </td>
            <td class="p-2.5">
                <input type="text" class="w-full bg-gray-50 border rounded px-2 py-1 text-xs focus:outline-none" value="${item.OBSERVACIONES || ''}" 
                       onchange="updateObs(${index}, this.value)" placeholder="Añadir nota...">
            </td>
        `;
        tbody.appendChild(tr);
    });
}

async function updateStatus(idx, val) {
    repartoData[idx].ESTADO_ENTREGA = val;
    if (val === 'PENDIENTE_TRANSF') repartoData[idx].FORMA_PAGO = 'TRANSFERENCIA';
    renderAll();
    await updateCloudRow(idx);
}

async function updatePayment(idx, val) {
    repartoData[idx].FORMA_PAGO = val;
    await updateCloudRow(idx);
}

async function updateObs(idx, val) {
    repartoData[idx].OBSERVACIONES = val;
    await updateCloudRow(idx);
}

// Función para actualizar el importe (Rechazo Parcial)
async function updateImporte(idx, val) {
    const nuevoMonto = cleanNumber(val);
    repartoData[idx].IMPORTE = nuevoMonto;
    calculateTotals();
    await updateCloudRow(idx);
}

// Sincronización completa con la nube incluyendo el importe modificado
async function updateCloudRow(index) {
    const item = repartoData[index];
    try {
        await fetch(SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify({
                action: 'updateRow',
                rowIndex: item.rowIndex,
                estadoEntrega: item.ESTADO_ENTREGA,
                formaPago: item.FORMA_PAGO,
                observacion: item.OBSERVACIONES,
                importe: item.IMPORTE // Enviamos el importe actualizado por si hubo rechazo parcial
            })
        });
    } catch(e) {
        console.error("Error al sincronizar en tiempo real", e);
    }
}

// Cálculo de métricas separado (Efectivo/Transf/Cheque vs Cta. Cte.)
function calculateTotals() {
    let totalCarga = 0, cobradoEfectivo = 0, totalCreditos = 0, pendTransf = 0, rechazado = 0;
    const clientesUnicos = new Set();

    repartoData.forEach(item => {
        const imp = cleanNumber(item.IMPORTE);
        totalCarga += imp;
        if (item.COD_CLIENTE) clientesUnicos.add(item.COD_CLIENTE);

        // Si es cuenta corriente, va a créditos y NO suma al efectivo cobrado
        if (item.FORMA_PAGO === 'CUENTA CORRIENTE') {
            totalCreditos += imp;
        } else if (item.ESTADO_ENTREGA === 'ENTREGADO') {
            cobradoEfectivo += imp; // Efectivo, Transferencia o Cheque cobrado
        }

        if (item.ESTADO_ENTREGA === 'PENDIENTE_TRANSF') pendTransf += imp;
        else if (item.ESTADO_ENTREGA === 'RECHAZADO_TOTAL' || item.ESTADO_ENTREGA === 'RECHAZADO_PARCIAL') rechazado += imp;
    });

    document.getElementById('totalCarga').textContent = formatMoney(totalCarga);
    document.getElementById('totalCobrado').textContent = formatMoney(cobradoEfectivo);
    document.getElementById('totalCreditos').textContent = formatMoney(totalCreditos);
    document.getElementById('totalPendTransf').textContent = formatMoney(pendTransf);
    document.getElementById('totalRechazado').textContent = formatMoney(rechazado);

    document.getElementById('cantClientes').textContent = `${clientesUnicos.size} Únicos`;
    document.getElementById('cantFacturas').textContent = `${repartoData.length} Facturas`;
}

function filterData() {
    const q = document.getElementById('searchDriverInput').value.toLowerCase();
    const filtered = repartoData.filter(i => 
        String(i.COD_CLIENTE).toLowerCase().includes(q) ||
        String(i.RAZON_SOCIAL).toLowerCase().includes(q) ||
        String(i.DIRECCION).toLowerCase().includes(q) ||
        String(i.COMPROBANTE).toLowerCase().includes(q)
    );
    renderAll(filtered);
}

async function loadHDRDataCloud(hdr) {
    try {
        const res = await fetch(`${SCRIPT_URL}?action=getHDR&hdr=${hdr}`);
        const json = await res.json();
        if(json.status === 'success') {
            repartoData = mapCloudData(json.data);
            if(repartoData.length > 0) {
                document.getElementById('hdrChofer').value = repartoData[0].CHOFER || '';
                document.getElementById('hdrZona').value = repartoData[0].ZONA || '';
                document.getElementById('hdrPatente').value = repartoData[0].PATENTE || '';
                renderAll();
            } else {
                // 🛑 Si no existe la HDR, avisa, limpia y devuelve al menú principal
                alert("⚠️ El número de HDR ingresado (" + hdr + ") no existe o es incorrecto.");
                location.reload(); // Vuelve a la pantalla de selección de perfil
            }
        }
    } catch(e) {
        alert("Error al conectar con el servidor: " + e);
        location.reload();
    }
}