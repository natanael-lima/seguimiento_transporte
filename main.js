/* ============================================================================
   1. VARIABLES GLOBALES Y ESTADO
   ============================================================================ */

let currentRole = '';        // 'admin' | 'driver'
let repartoData = [];        // Array principal con todas las filas de la HDR activa
let adminTab = 'create';     // 'create' | 'monitor'
let cobroIndex = -1;         // Índice de la fila que se está editando en el modal de cobro


/* ============================================================================
   2. UTILIDADES GENERALES
   ============================================================================ */

// Inicializa la fecha por defecto en el campo #hdrFecha al cargar el DOM
document.addEventListener("DOMContentLoaded", function() {
    const dateField = document.getElementById('hdrFecha');
    if (dateField) dateField.value = new Date().toISOString().slice(0, 10);
});

// Devuelve la fecha y hora actual formateada como "DD/MM/YYYY HH:MM"
function getFormattedDateTime() {
    const now = new Date();
    const day = String(now.getDate()).padStart(2, '0');
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const year = now.getFullYear();
    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    return `${day}/${month}/${year} ${hours}:${minutes}`;
}

// Setea el timestamp inicial en el panel superior
document.getElementById('hdrTimestamp').value = getFormattedDateTime();

// Convierte cualquier valor (string con $, puntos, comas) a número flotante
function cleanNumber(val) {
    if (typeof val === 'number') return val;
    if (!val) return 0;
    let str = String(val).replace(/\$/g, '').trim();
    if (str.includes('.') && str.includes(',')) str = str.replace(/\./g, '').replace(',', '.');
    else if (str.includes(',')) str = str.replace(',', '.');
    return parseFloat(str) || 0;
}

// Formatea un número como moneda argentina ($ 12.345,67)
function formatMoney(num) {
    return '$ ' + num.toLocaleString('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}


/* ============================================================================
   3. AUTENTICACIÓN Y SELECCIÓN DE ROL
   ============================================================================ */

// Muestra/oculta el cuadro de contraseña al hacer click en "Administrador"
function requestAdminAccess() {
    const passwordBox = document.getElementById('adminPasswordBox');
    passwordBox.classList.toggle('hidden');
    if (!passwordBox.classList.contains('hidden')) {
        document.getElementById('adminPasswordInput').focus();
    }
}

// Valida la clave de admin y entra al sistema
function verifyAdminPassword() {
    const pass = document.getElementById('adminPasswordInput').value.trim();
    if (pass === "admin") {
        selectRole('admin');
    } else {
        alert("Contraseña incorrecta.");
        document.getElementById('adminPasswordInput').value = '';
    }
}

// Selecciona el rol (admin o driver) y muestra la UI correspondiente
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


/* ============================================================================
   4. NAVEGACIÓN ENTRE PESTAÑAS DEL ADMIN
   ============================================================================ */

// Cambia entre la pestaña "Cargar Nueva HDR" y "Monitorear Repartos Activos"
async function switchAdminTab(tab) {
    adminTab = tab;
    const btnCreate = document.getElementById('btnTabCreate');
    const btnMonitor = document.getElementById('btnTabMonitor');
    const createControls = document.getElementById('adminCreateControls');
    const monitorBox = document.getElementById('monitorSelectorBox');
    const searchContainer = document.getElementById('generalSearchContainer');

    // En "create" los campos del header son editables; en "monitor" son solo lectura
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


/* ============================================================================
   5. CARGA DE EXCEL (ADMIN)
   ----------------------------------------------------------------------------
   Lee el archivo Excel/CSV, mapea las columnas y arma el array repartoData.
   ============================================================================ */

document.getElementById('fileInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(evt) {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws, { defval: '' });

        // Toma los datos del header del formulario
        const hdrInput = document.getElementById('hdrNumero').value.trim() || 'S/N';
        const choferInput = document.getElementById('hdrChofer').value.trim() || 'S/D';
        const zonaInput = document.getElementById('hdrZona').value.trim() || 'S/D';
        const patenteInput = document.getElementById('hdrPatente').value.trim() || 'S/D';

        // Fecha de hoy en formato DD/MM/YYYY
        const now = new Date();
        const day = String(now.getDate()).padStart(2, '0');
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const year = now.getFullYear();
        const fechaHoy = `${day}/${month}/${year}`;

        // Actualiza el timestamp del panel superior
        const timestampEl = document.getElementById('hdrTimestamp');
        if (timestampEl) {
            const hours = String(now.getHours()).padStart(2, '0');
            const minutes = String(now.getMinutes()).padStart(2, '0');
            timestampEl.value = `${fechaHoy} ${hours}:${minutes}`;
        }

        // Mapea cada fila del Excel a la estructura interna
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
                COBRADO: 0,                    // ⬅️ NUEVO: inicia en 0
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


/* ============================================================================
   6. REGISTRO DE HDR EN LA NUBE (ADMIN)
   ----------------------------------------------------------------------------
   Verifica si la HDR ya existe y luego envía todos los registros al backend.
   ============================================================================ */

async function registerHDRCloud() {
    if (repartoData.length === 0) return alert("No hay datos cargados para registrar.");

    const hdrVal = document.getElementById('hdrNumero').value.trim();
    const choferVal = document.getElementById('hdrChofer').value.trim();
    const zonaVal = document.getElementById('hdrZona').value.trim();
    const patenteVal = document.getElementById('hdrPatente').value.trim();

    if (!hdrVal) return alert("Por favor, ingrese el número de HDR antes de registrar.");

    try {
        // 1. Verifica si la HDR ya existe en la base de datos
        const checkRes = await fetch(`${SCRIPT_URL}?action=getHDR&hdr=${hdrVal}`);
        const checkJson = await checkRes.json();

        if (checkJson.status === 'success' && checkJson.data && checkJson.data.length > 0) {
            const confirmar = confirm(`⚠️ ¡Atención! La Hoja de Ruta N° ${hdrVal} ya se encuentra registrada en el sistema.\n\n¿Desea agregar estos registros de todas formas?`);
            if (!confirmar) return;
        }

        // 2. Asegura que cada fila tenga los datos del header
        repartoData.forEach(i => {
            i.HDR = hdrVal;
            i.CHOFER = choferVal;
            i.ZONA = zonaVal;
            i.PATENTE = patenteVal;
        });

        // 3. Envía todo al backend
        const res = await fetch(SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify({ action: 'saveAll', rows: repartoData })
        });
        const json = await res.json();

        if (json.status === 'success') {
            alert("¡Hoja de Ruta " + hdrVal + " registrada en Google Sheets con éxito!");
            switchAdminTab('monitor');
        } else {
            alert("Error al registrar: " + json.message);
        }

    } catch (e) {
        alert("Error de conexión al registrar: " + e);
    }
}


/* ============================================================================
   7. CONSULTA DE HDR (MONITOR / CHOFER)
   ============================================================================ */

// Admin: consulta una HDR desde el input del monitor
async function searchMonitorHDR() {
    const hdr = document.getElementById('inputMonitorHDR').value.trim();
    if (!hdr) return alert("Ingrese un número de HDR para consultar.");

    try {
        const res = await fetch(`${SCRIPT_URL}?action=getHDR&hdr=${hdr}`);
        const json = await res.json();
        if (json.status === 'success') {
            repartoData = mapCloudData(json.data);
            if (repartoData.length > 0) {
                document.getElementById('hdrNumero').value = repartoData[0].HDR || hdr;
                document.getElementById('hdrChofer').value = repartoData[0].CHOFER || '';
                document.getElementById('hdrZona').value = repartoData[0].ZONA || '';
                document.getElementById('hdrPatente').value = repartoData[0].PATENTE || '';
            } else {
                alert("No se encontraron registros para la HDR " + hdr);
            }
            renderAll();
        }
    } catch (e) {
        alert("Error al consultar ruta: " + e);
    }
}

// Chofer: toma la HDR del modal de login y la carga
function loadDriverHDR() {
    const hdr = document.getElementById('inputHdrLogin').value.trim();
    if (!hdr) return alert("Ingrese un número de HDR válido.");
    document.getElementById('driverModal').classList.add('hidden');
    document.getElementById('hdrNumero').value = hdr;
    loadHDRDataCloud(hdr);
}

// Carga la HDR desde la nube (usado por chofer y admin)
async function loadHDRDataCloud(hdr) {
    try {
        const res = await fetch(`${SCRIPT_URL}?action=getHDR&hdr=${hdr}`);
        const json = await res.json();
        if (json.status === 'success') {
            repartoData = mapCloudData(json.data);
            if (repartoData.length > 0) {
                document.getElementById('hdrChofer').value = repartoData[0].CHOFER || '';
                document.getElementById('hdrZona').value = repartoData[0].ZONA || '';
                document.getElementById('hdrPatente').value = repartoData[0].PATENTE || '';
                renderAll();
            } else {
                alert("⚠️ El número de HDR ingresado (" + hdr + ") no existe o es incorrecto.");
                location.reload();
            }
        }
    } catch (e) {
        alert("Error al conectar con el servidor: " + e);
        location.reload();
    }
}


/* ============================================================================
   8. MAPEO DE DATOS DESDE LA NUBE
   ----------------------------------------------------------------------------
   Convierte las filas crudas del backend a la estructura interna.
   ============================================================================ */

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
            COBRADO: cleanNumber(item.COBRADO || 0),   // ⬅️ NUEVO
            ESTADO_ENTREGA: item.ESTADO_ENTREGA || item['ESTADO ENTREGA'] || item.ESTADO || 'PENDIENTE',
            FORMA_PAGO: item.FORMA_PAGO || item['FORMA DE PAGO'] || 'EFECTIVO',
            OBSERVACIONES: item.OBSERVACIONES || '',
            FECHA: item.FECHA || ''
        };
    });
}


/* ============================================================================
   9. SINCRONIZACIÓN DE FILAS CON LA NUBE
   ----------------------------------------------------------------------------
   Envía los cambios (estado, forma de pago, observación, cobrado) al backend.
   ============================================================================ */

async function updateCloudRow(index) {
    const item = repartoData[index];

    if (!item) {
        console.error("updateCloudRow: item undefined en index", index);
        return;
    }
    try {
        await fetch(SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify({
                action: 'updateRow',
                rowIndex: item.rowIndex,
                estadoEntrega: item.ESTADO_ENTREGA,
                formaPago: item.FORMA_PAGO,
                observacion: item.OBSERVACIONES,
                cobrado: item.COBRADO     // ⬅️ NUEVO
            })
        });
    } catch (e) {
        console.error("Error al sincronizar en tiempo real", e);
    }
}

// Actualiza la observación de una fila y sincroniza con la nube
async function updateObs(idx, val) {
    repartoData[idx].OBSERVACIONES = val;
    await updateCloudRowByRowIndex(repartoData[idx].rowIndex);
}


/* ============================================================================
   10. LÓGICA DE ESTADOS (DEDUCCIÓN AUTOMÁTICA)
   ----------------------------------------------------------------------------
   El estado NO se elige manualmente: se deduce según COBRADO vs IMPORTE.
   ============================================================================ */

function deducirEstado(item) {
    const cobrado = cleanNumber(item.COBRADO);
    const importe = cleanNumber(item.IMPORTE);
    const obs = (item.OBSERVACIONES || '').toLowerCase().trim();

    // 🟣 REASIGNADO: no cobró y avisó que vuelve / reasigna
    if (cobrado === 0 && (
        obs.includes('rea') ||
        obs.includes('reasignado') ||
        obs.includes('reasignados')
    )) {
        return 'REASIGNADO';
    }

    // 🔴 RECHAZO TOTAL: no cobró y avisó que rechaza
    if (cobrado === 0 && (
        obs.includes('rechaz') ||
        obs.includes('no recib') ||
        obs.includes('no quiso') ||
        obs.includes('no acept')
    )) {
        return 'RECHAZO_TOTAL';
    }

    // 🟢 ENTREGADO: cobró todo
    if (cobrado >= importe && cobrado > 0) return 'ENTREGADO';

    // 🟠 RECHAZO PARCIAL: cobró algo pero menos
    if (cobrado > 0 && cobrado < importe) return 'RECHAZO_PARCIAL';

    // 🟡 PENDIENTE: no tocó nada
    return 'PENDIENTE';
}


/* ============================================================================
   11. RENDERIZADO (MOBILE CARDS + DESKTOP TABLE)
   ============================================================================ */

// Renderiza ambas vistas + recalcula métricas
function renderAll(filteredData = null) {
    const list = filteredData || repartoData;
    renderMobileCards(list);
    renderDesktopTable(list);
    calculateTotals();
}

// Vista móvil: tarjetas
// Vista móvil: tarjetas (versión resumida)
function renderMobileCards(list) {
    const container = document.getElementById('cardsContainer');
    container.innerHTML = '';

    if (list.length === 0) {
        container.innerHTML = '<div class="text-center py-8 bg-white rounded-xl border text-gray-400 text-xs">No hay registros para mostrar.</div>';
        return;
    }

    list.forEach((item, index) => {
        const card = document.createElement('div');
        card.className = `bg-white p-3 rounded-xl border border-gray-200 shadow-xs flex flex-col gap-2.5 ${getStatusBorderClass(item.ESTADO_ENTREGA)}`;

        card.innerHTML = `
            <!-- Nombre + Importe -->
            <div class="flex justify-between items-start gap-2">
                <div class="text-sm font-bold text-slate-900 leading-tight">${item.RAZON_SOCIAL}</div>
                <span class="text-sm font-extrabold text-slate-900 whitespace-nowrap">${formatMoney(item.IMPORTE)}</span>
            </div>

            <!-- Dirección -->
            <div class="text-[11px] text-gray-500 leading-tight">
                <i class="fa-solid fa-location-dot text-gray-400 mr-1"></i>${item.DIRECCION}
            </div>

            <!-- Estado + Cobrado -->
            <div class="flex items-center justify-between gap-2">
                <span class="px-2 py-1 rounded-lg border text-[10px] font-bold inline-flex items-center gap-1 ${getStatusSelectClass(item.ESTADO_ENTREGA)}">
                    ${getEstadoLabel(item.ESTADO_ENTREGA)}
                </span>
                <span class="text-xs font-bold text-emerald-700">${formatMoney(item.COBRADO || 0)}</span>
            </div>

            <!-- Botón Registrar -->
           <button onclick="openCobroModalByRowIndex(${item.rowIndex})" 
                class="w-full bg-white hover:bg-slate-100 border border-slate-300 hover:border-slate-400 text-slate-900 py-2 rounded-lg text-xs font-bold inline-flex items-center justify-center gap-2 transition-all duration-200 ease-out hover:scale-[1.02] active:scale-[0.98] shadow-sm hover:shadow-md">
                <i class="fa-solid fa-clipboard-check text-slate-500"></i> Registrar
            </button>
        `;
        container.appendChild(card);
    });
}

// Vista escritorio: tabla
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
            <td class="p-2.5 font-bold text-emerald-700">${formatMoney(item.COBRADO || 0)}</td>

            <!-- COLUMNA 8: Estado (solo indicador visual) -->
            <td class="p-2.5">
                <span class="px-2 py-1 rounded-lg border text-xs font-bold inline-flex items-center gap-1 ${getStatusSelectClass(item.ESTADO_ENTREGA)}">
                    ${getEstadoLabel(item.ESTADO_ENTREGA)}
                </span>
            </td>

            <!-- COLUMNA 9: Botón Cobrar (reemplaza al select) -->
           <td class="p-2.5">
                <button onclick="openCobroModalByRowIndex(${item.rowIndex})" 
                    class="bg-slate-100 hover:bg-slate-200 border border-slate-300 text-slate-800 px-3.5 py-1.5 rounded-lg text-xs font-bold inline-flex items-center gap-2 transition-all duration-200 ease-out hover:scale-[1.03] active:scale-[0.98] shadow-sm hover:shadow-md whitespace-nowrap">
                    <i class="fa-solid fa-clipboard-check text-slate-600"></i> Registrar
                </button>
            </td>

            <!-- COLUMNA 10: Observación -->
            <td class="p-2.5">
                <input type="text" class="w-full bg-gray-50 border rounded px-2 py-1 text-xs focus:outline-none" value="${item.OBSERVACIONES || ''}" 
                       onchange="updateObs(${index}, this.value)" placeholder="Añadir nota...">
            </td>
        `;
        tbody.appendChild(tr);
    });
}


/* ============================================================================
   12. ESTILOS VISUALES DE ESTADOS (COLORES, BORDES, LABELS)
   ============================================================================ */
/* ============================================================================
   12. ESTILOS VISUALES DE ESTADOS (COLORES, BORDES, LABELS)
   ============================================================================ */
function getStatusBorderClass(estado) {
    if (estado === 'ENTREGADO')         return 'border-l-4 border-l-green-600';
    if (estado === 'RECHAZO_PARCIAL')   return 'border-l-4 border-l-amber-500'; // Naranja bien diferenciado
    if (estado === 'RECHAZO_TOTAL')     return 'border-l-4 border-l-red-600';
    if (estado === 'REASIGNADO')        return 'border-l-4 border-l-purple-600';
    return 'border-l-4 border-l-slate-400'; // PENDIENTE en Gris neutro (elimina la confusión con el ámbar)
}

function getStatusSelectClass(estado) {
    if (estado === 'ENTREGADO')         return 'bg-green-50 text-green-700 border-green-300';
    if (estado === 'RECHAZO_PARCIAL')   return 'bg-amber-50 text-amber-700 border-amber-300';
    if (estado === 'RECHAZO_TOTAL')     return 'bg-red-50 text-red-700 border-red-300';
    if (estado === 'REASIGNADO')        return 'bg-purple-50 text-purple-700 border-purple-300';
    return 'bg-slate-100 text-slate-700 border-slate-300'; // PENDIENTE limpio y neutro
}

function getEstadoLabel(estado) {
    switch (estado) {
        case 'ENTREGADO':       return '🟢 ENTREGADO';
        case 'RECHAZO_PARCIAL': return '🟠 RECHAZO PARCIAL';
        case 'RECHAZO_TOTAL':   return '🔴 RECHAZO TOTAL';
        case 'REASIGNADO':      return '🟣 REASIGNADO';
        default:                return '⚪ PENDIENTE'; // Cambia al círculo blanco/gris
    }
}

/* ============================================================================
   13. MODAL DE COBRO
   ----------------------------------------------------------------------------
   Abre un modal donde el chofer ingresa el monto cobrado, la forma de pago
   y la observación. El estado se deduce automáticamente.
   ============================================================================ */

// Abre el modal buscando por rowIndex real del Sheet (robusto ante filtros)
function openCobroModalByRowIndex(rowIndex) {
    const index = repartoData.findIndex(i => i.rowIndex === rowIndex);
    if (index < 0) {
        alert("⚠️ No se encontró el registro. Recargá la HDR.");
        return;
    }
    openCobroModal(index);
}
// Abre el modal con los datos de la fila seleccionada
/* ============================================================================
   13. MODAL DE COBRO
   ============================================================================ */

function openCobroModal(index) {
    cobroIndex = index;
    const item = repartoData[index];
    if (!item) return;

    // Datos del cliente
    document.getElementById('modalCliente').textContent = item.RAZON_SOCIAL;
    document.getElementById('modalDireccion').textContent = item.DIRECCION;
    document.getElementById('modalComprobante').textContent = item.COMPROBANTE;
    document.getElementById('modalImporte').textContent = formatMoney(item.IMPORTE);
    document.getElementById('modalCobrado').value = item.COBRADO || '';
    document.getElementById('modalFormaPago').value = item.FORMA_PAGO || 'EFECTIVO';
    document.getElementById('modalObs').value = item.OBSERVACIONES || '';

    // Resetea el checklist a "Ninguno"
    document.querySelector('input[name="resultadoVisita"][value="ninguno"]').checked = true;

    // Habilita los inputs por defecto
    habilitarInputsCobro(true);

    updateEstadoPreview();
    document.getElementById('cobroModal').classList.remove('hidden');
    document.getElementById('cobroModal').classList.add('flex');
}

function closeCobroModal() {
    document.getElementById('cobroModal').classList.add('hidden');
    document.getElementById('cobroModal').classList.remove('flex');
    cobroIndex = -1;
}

// Habilita o deshabilita los inputs de monto/pago según el checklist
function habilitarInputsCobro(habilitar) {
    const inputCobrado = document.getElementById('modalCobrado');
    const selectPago = document.getElementById('modalFormaPago');
    const btnCobrarTotal = document.getElementById('btnCobrarTotal');

    inputCobrado.disabled = !habilitar;
    selectPago.disabled = !habilitar;
    btnCobrarTotal.disabled = !habilitar;

    if (habilitar) {
        inputCobrado.classList.remove('bg-gray-100', 'cursor-not-allowed', 'opacity-60');
        selectPago.classList.remove('bg-gray-100', 'cursor-not-allowed', 'opacity-60');
        btnCobrarTotal.classList.remove('opacity-50', 'cursor-not-allowed');
    } else {
        inputCobrado.classList.add('bg-gray-100', 'cursor-not-allowed', 'opacity-60');
        selectPago.classList.add('bg-gray-100', 'cursor-not-allowed', 'opacity-60');
        btnCobrarTotal.classList.add('opacity-50', 'cursor-not-allowed');
    }
}

// Aplica el resultado rápido seleccionado en el checklist
function aplicarResultadoRapido() {
    const seleccionado = document.querySelector('input[name="resultadoVisita"]:checked')?.value;

    const inputCobrado = document.getElementById('modalCobrado');
    const obsInput = document.getElementById('modalObs');

    if (seleccionado === 'rechazo_total') {
        inputCobrado.value = 0;
        obsInput.value = 'RECHAZO TOTAL';
        habilitarInputsCobro(false);
    } else if (seleccionado === 'reasignado') {
        inputCobrado.value = 0;
        obsInput.value = 'REASIGNADO - vuelve';
        habilitarInputsCobro(false);
    } else {
        // "Ninguno" → vuelve a modo normal
        habilitarInputsCobro(true);
        // No tocamos el monto ni la observación (el chofer decide)
    }

    updateEstadoPreview();
}

// Botón "Cobrar Total"
function setCobroTotal() {
    const item = repartoData[cobroIndex];
    if (item) {
        document.getElementById('modalCobrado').value = item.IMPORTE;
        updateEstadoPreview();
    }
}

// Vista previa en vivo del estado
function updateEstadoPreview() {
    const cobrado = cleanNumber(document.getElementById('modalCobrado').value);
    const formaPago = document.getElementById('modalFormaPago').value;
    const obs = document.getElementById('modalObs').value;
    const importe = repartoData[cobroIndex]?.IMPORTE || 0;

    const estado = deducirEstado({
        COBRADO: cobrado,
        IMPORTE: importe,
        FORMA_PAGO: formaPago,
        OBSERVACIONES: obs
    });

    const preview = document.getElementById('modalEstadoPreview');
    preview.textContent = 'Estado: ' + getEstadoLabel(estado);
    preview.className = 'text-center py-2 rounded-lg border text-xs font-bold ' + getStatusSelectClass(estado);
}

// Escuchar cambios dentro del modal para actualizar la vista previa
document.addEventListener('input', e => {
    if (['modalCobrado', 'modalFormaPago', 'modalObs'].includes(e.target.id)) {
        updateEstadoPreview();
    }
});

// Guardar el cobro
async function guardarCobro() {
    if (cobroIndex < 0) return;
    const item = repartoData[cobroIndex];
    if (!item) {
        alert("⚠️ Registro no encontrado. Recargá la HDR.");
        closeCobroModal();
        return;
    }

    item.COBRADO = cleanNumber(document.getElementById('modalCobrado').value);
    item.FORMA_PAGO = document.getElementById('modalFormaPago').value;
    item.OBSERVACIONES = document.getElementById('modalObs').value;
    item.ESTADO_ENTREGA = deducirEstado(item);

    // Validación: rechazo parcial requiere observación
    if (item.ESTADO_ENTREGA === 'RECHAZO_PARCIAL' && !item.OBSERVACIONES.trim()) {
        alert('⚠️ Hay una diferencia de cobro. Es obligatorio dejar una observación.');
        return;
    }

    // Guardamos el rowIndex antes de cerrar (closeCobroModal resetea cobroIndex)
    const rowIndexToSync = item.rowIndex;

    closeCobroModal();
    renderAll();
    await updateCloudRowByRowIndex(rowIndexToSync);
}
async function updateCloudRowByRowIndex(rowIndex) {
    const item = repartoData.find(i => i.rowIndex === rowIndex);
    if (!item) {
        console.error("No se encontró item con rowIndex", rowIndex);
        return;
    }
    try {
        await fetch(SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify({
                action: 'updateRow',
                rowIndex: item.rowIndex,
                estadoEntrega: item.ESTADO_ENTREGA,
                formaPago: item.FORMA_PAGO,
                observacion: item.OBSERVACIONES,
                cobrado: item.COBRADO
            })
        });
    } catch (e) {
        console.error("Error al sincronizar en tiempo real", e);
    }
}


/* ============================================================================
   14. CÁLCULO DE MÉTRICAS / TOTALES
   ----------------------------------------------------------------------------
   Suma los totales para las tarjetas superiores.
   ============================================================================ */

function calculateTotals() {
    let totalCarga = 0, cobradoEfectivo = 0, totalCreditos = 0, rechazado = 0;
    const clientesUnicos = new Set();

    repartoData.forEach(item => {
        const imp = cleanNumber(item.IMPORTE);
        const cob = cleanNumber(item.COBRADO);
        totalCarga += imp;
        if (item.COD_CLIENTE) clientesUnicos.add(item.COD_CLIENTE);

        // 💳 Si es Cuenta Corriente, va a Créditos (sin importar el estado)
        if (item.FORMA_PAGO === 'CUENTA CORRIENTE') {
            totalCreditos += cob;
        }  // 💵 Si cobró algo (ENTREGADO o RECHAZO_PARCIAL), suma a Cobrado
        else if (item.ESTADO_ENTREGA === 'ENTREGADO' || item.ESTADO_ENTREGA === 'RECHAZO_PARCIAL') {
            cobradoEfectivo += cob;
        }
        // 🚫 Rechazado: 
        //   - RECHAZO_PARCIAL → solo la diferencia (lo que NO cobró)
        //   - RECHAZO_TOTAL   → todo el importe
        if (item.ESTADO_ENTREGA === 'RECHAZO_PARCIAL') {
            rechazado += (imp - cob);
        } else if (item.ESTADO_ENTREGA === 'RECHAZO_TOTAL') {
            rechazado += imp;
        }
    });

    document.getElementById('totalCarga').textContent = formatMoney(totalCarga);
    document.getElementById('totalCobrado').textContent = formatMoney(cobradoEfectivo);
    document.getElementById('totalCreditos').textContent = formatMoney(totalCreditos);
    document.getElementById('totalRechazado').textContent = formatMoney(rechazado);

    document.getElementById('cantClientes').textContent = `${clientesUnicos.size} Únicos`;
    document.getElementById('cantFacturas').textContent = `${repartoData.length} Facturas`;
}

/* ============================================================================
   15. BÚSQUEDA / FILTROS
   ============================================================================ */

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


/* ============================================================================
   16. EXPORTACIÓN A EXCEL (ADMIN)
   ----------------------------------------------------------------------------
   Exporta la HDR que el admin está monitoreando a un archivo .xlsx.
   ============================================================================ */

function exportMonitorExcel() {
    if (repartoData.length === 0) return alert("No hay datos de ruta cargados para exportar.");

    const hdrNum = document.getElementById('hdrNumero').value || 'REPORTE';
    const chofer = document.getElementById('hdrChofer').value || 'S_D';

    // Arma las filas para el Excel
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
        "COBRADO": i.COBRADO,
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