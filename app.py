<!DOCTYPE html>
<html>
<head>
    <title>Panel de control</title>
    <style>
        body { font-family: Arial; margin: 20px; }
        .pestanas { margin-bottom: 20px; }
        .pestana { display: inline-block; padding: 10px; background: #ddd; cursor: pointer; margin-right: 5px; }
        .pestana.activa { background: #007bff; color: white; }
        table { border-collapse: collapse; width: 100%; overflow-x: auto; display: block; }
        th, td { border: 1px solid #ccc; padding: 8px; text-align: left; vertical-align: top; }
        th { background: #f2f2f2; }
        textarea { width: 100%; box-sizing: border-box; }
        .guardar { background: #28a745; color: white; border: none; padding: 5px 10px; cursor: pointer; }
        .acciones { margin: 20px 0; }
        button { padding: 10px; margin-right: 10px; }
    </style>
</head>
<body>
    <h2>Bienvenido, {{ usuario }}</h2>
    <div class="acciones">
        <button id="btnExportar">📥 Exportar Excel</button>
        <button id="btnSubir" {% if usuario != 'ANA JULCA' %}disabled{% endif %}>📂 Subir Excel nuevo (solo admin)</button>
        <button id="btnPausa" {% if usuario != 'ANA JULCA' %}disabled{% endif %}>⏸️ Pausar / Reanudar</button>
        <span id="estadoPausa" style="margin-left: 20px;"></span>
        <a href="/logout" style="margin-left: 20px;">Cerrar sesión</a>
    </div>
    <div class="pestanas">
        <div id="pestana_SCD" class="pestana" data-hoja="SCD-2026">SCD-2026</div>
        <div id="pestana_SAF" class="pestana" data-hoja="SAF-2026">SAF-2026</div>
    </div>
    <div id="contenedor_tabla">
        <p>Cargando datos...</p>
    </div>

    <script>
        let hojaActual = 'SCD-2026';
        let datosOriginales = [];

        async function cargarDatos() {
            const res = await fetch(`/get_datos?hoja=${hojaActual}`);
            const json = await res.json();
            datosOriginales = json.datos;
            const columnas = json.columnas;
            if (!columnas || columnas.length === 0) {
                document.getElementById('contenedor_tabla').innerHTML = '<p>No hay filas para ti en esta hoja.</p>';
                return;
            }
            let html = '<table><thead><tr>';
            columnas.forEach(col => {
                html += `<th>${escapeHtml(col)}</th>`;
            });
            html += '<th>Acción</th></tr></thead><tbody>';
            for (let i = 0; i < datosOriginales.length; i++) {
                const fila = datosOriginales[i];
                const rowid = fila.rowid;
                html += `<tr id="fila_${i}">`;
                for (let col of columnas) {
                    if (col === 'rowid') continue; // no mostrar la columna rowid
                    const valor = fila[col] || '';
                    html += `<td><textarea rows="2" style="width:100%" data-col="${col}" data-fila="${i}">${escapeHtml(valor)}</textarea></td>`;
                }
                html += `<td><button class="guardar" data-fila="${i}" data-rowid="${rowid}">Guardar cambios</button></td>`;
                html += '</tr>';
            }
            html += '</tbody></table>';
            document.getElementById('contenedor_tabla').innerHTML = html;
            // Asignar eventos a los botones guardar
            document.querySelectorAll('.guardar').forEach(btn => {
                btn.addEventListener('click', async (e) => {
                    const filaIdx = btn.dataset.fila;
                    const rowid = btn.dataset.rowid;
                    const campos = {};
                    const textareas = document.querySelectorAll(`#fila_${filaIdx} textarea`);
                    textareas.forEach(ta => {
                        const col = ta.dataset.col;
                        campos[col] = ta.value;
                    });
                    const respuesta = await fetch('/guardar_fila', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ hoja: hojaActual, rowid: rowid, campos: campos })
                    });
                    const result = await respuesta.json();
                    if (result.ok) {
                        alert('Guardado exitoso');
                        // Actualizar datosOriginales
                        const filaOriginal = datosOriginales[filaIdx];
                        for (let col in campos) {
                            filaOriginal[col] = campos[col];
                        }
                    } else {
                        alert('Error: ' + (result.error || 'No se pudo guardar'));
                    }
                });
            });
        }

        function escapeHtml(text) {
            if (!text) return '';
            return text.replace(/[&<>]/g, function(m) {
                if (m === '&') return '&amp;';
                if (m === '<') return '&lt;';
                if (m === '>') return '&gt;';
                return m;
            });
        }

        document.getElementById('pestana_SCD').addEventListener('click', () => {
            hojaActual = 'SCD-2026';
            document.querySelectorAll('.pestana').forEach(p => p.classList.remove('activa'));
            document.getElementById('pestana_SCD').classList.add('activa');
            cargarDatos();
        });
        document.getElementById('pestana_SAF').addEventListener('click', () => {
            hojaActual = 'SAF-2026';
            document.querySelectorAll('.pestana').forEach(p => p.classList.remove('activa'));
            document.getElementById('pestana_SAF').classList.add('activa');
            cargarDatos();
        });
        document.getElementById('btnExportar').addEventListener('click', () => {
            window.location.href = '/exportar';
        });
        document.getElementById('btnSubir').addEventListener('click', () => {
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = '.xlsx, .xls';
            input.onchange = async (e) => {
                const file = e.target.files[0];
                const formData = new FormData();
                formData.append('archivo', file);
                const res = await fetch('/subir_excel', { method: 'POST', body: formData });
                const result = await res.json();
                if (result.ok) {
                    alert('Excel subido correctamente. Recargando datos...');
                    cargarDatos();
                } else {
                    alert('Error: ' + result.error);
                }
            };
            input.click();
        });
        document.getElementById('btnPausa').addEventListener('click', async () => {
            const res = await fetch('/toggle_pausa', { method: 'POST' });
            const result = await res.json();
            if (result.pausa !== undefined) {
                actualizarEstadoPausa(result.pausa);
            } else {
                alert('Error al cambiar estado');
            }
        });

        async function actualizarEstadoPausa(pausa) {
            const span = document.getElementById('estadoPausa');
            if (pausa) {
                span.innerHTML = '🔴 SISTEMA EN PAUSA - No se pueden editar';
                span.style.color = 'red';
            } else {
                span.innerHTML = '🟢 SISTEMA ACTIVO - Ediciones permitidas';
                span.style.color = 'green';
            }
        }

        // Inicializar
        document.getElementById('pestana_SCD').classList.add('activa');
        cargarDatos();
        // Cargar estado de pausa (opcional)
        fetch('/toggle_pausa', { method: 'GET' }).catch(()=>{});
    </script>
</body>
</html>
