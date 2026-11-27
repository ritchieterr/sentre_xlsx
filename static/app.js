/**
 * PDF to XLSX Extractor - Frontend unificado (Trámite y Concentración)
 */

document.addEventListener('DOMContentLoaded', () => {
    const uploadZone = document.getElementById('uploadZone');
    const fileInput = document.getElementById('fileInput');
    const btnSelect = document.getElementById('btnSelect');
    const uploadContent = document.getElementById('uploadContent');
    const processStatus = document.getElementById('processStatus');
    const successContent = document.getElementById('successContent');
    const errorContent = document.getElementById('errorContent');
    const fileInfo = document.getElementById('fileInfo');
    const fileName = document.getElementById('fileName');
    const fileSize = document.getElementById('fileSize');
    const statusTitle = document.getElementById('statusTitle');
    const statusMessage = document.getElementById('statusMessage');
    const progressFill = document.getElementById('progressFill');
    const recordCount = document.getElementById('recordCount');
    const outputFileNameEl = document.getElementById('outputFileName');
    const errorMessage = document.getElementById('errorMessage');
    const btnDownload = document.getElementById('btnDownload');
    const btnReset = document.getElementById('btnReset');
    const btnRetry = document.getElementById('btnRetry');

    const MAX_FILE_SIZE_BYTES = 50 * 1024 * 1024; // 50 MB
    const uploadEndpoint = uploadZone?.dataset.endpoint || '/upload';
    const defaultFileName = uploadZone?.dataset.filename || 'inventario_extraido.xlsx';

    let currentBlob = null;
    let currentFileName = defaultFileName;
    let isProcessing = false;

    const STATES = {
        IDLE: 'idle',
        UPLOADING: 'uploading',
        READING: 'reading',
        PROCESSING: 'processing',
        EXPORTING: 'exporting',
        SUCCESS: 'success',
        ERROR: 'error',
    };

    const STATE_INFO = {
        [STATES.UPLOADING]: {
            title: 'Cargando archivo...',
            message: 'Subiendo el documento PDF al servidor',
            progress: 22,
        },
        [STATES.READING]: {
            title: 'Leyendo estructura del PDF...',
            message: 'Identificando tablas y filas de inventario',
            progress: 48,
        },
        [STATES.PROCESSING]: {
            title: 'Procesando registros...',
            message: 'Normalizando campos y unificando observaciones multilínea',
            progress: 74,
        },
        [STATES.EXPORTING]: {
            title: 'Generando plantilla PERGUB...',
            message: 'Preparando el archivo Excel (.xlsx) para descarga',
            progress: 94,
        },
    };

    function formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return `${parseFloat((bytes / Math.pow(k, i)).toFixed(2))} ${sizes[i]}`;
    }

    function showState(state) {
        uploadContent.classList.add('hidden');
        processStatus.classList.add('hidden');
        successContent.classList.add('hidden');
        errorContent.classList.add('hidden');
        uploadZone.classList.remove('is-busy');

        switch (state) {
            case STATES.IDLE:
                isProcessing = false;
                uploadContent.classList.remove('hidden');
                fileInfo.classList.add('hidden');
                progressFill.style.width = '0%';
                break;

            case STATES.UPLOADING:
            case STATES.READING:
            case STATES.PROCESSING:
            case STATES.EXPORTING: {
                isProcessing = true;
                uploadZone.classList.add('is-busy');
                processStatus.classList.remove('hidden');
                const info = STATE_INFO[state];
                statusTitle.textContent = info.title;
                statusMessage.textContent = info.message;
                progressFill.style.width = `${info.progress}%`;
                break;
            }

            case STATES.SUCCESS:
                isProcessing = false;
                uploadZone.classList.add('is-busy');
                successContent.classList.remove('hidden');
                fileInfo.classList.add('hidden');
                break;

            case STATES.ERROR:
                isProcessing = false;
                uploadZone.classList.add('is-busy');
                errorContent.classList.remove('hidden');
                fileInfo.classList.add('hidden');
                break;
        }
    }

    function resetUI() {
        currentBlob = null;
        currentFileName = defaultFileName;
        fileInput.value = '';
        showState(STATES.IDLE);
    }

    function triggerDownload() {
        if (!currentBlob) return;
        const url = URL.createObjectURL(currentBlob);
        const a = document.createElement('a');
        a.href = url;
        a.download = currentFileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    function sleep(ms) {
        return new Promise((resolve) => setTimeout(resolve, ms));
    }

    async function loadCounter() {
        try {
            const response = await fetch('/counter');
            if (!response.ok) return;
            const data = await response.json();
            const recordsEl = document.getElementById('globalCounter');
            const filesEl = document.getElementById('globalFilesCounter');
            if (recordsEl) {
                recordsEl.textContent = (data.total_records || 0).toLocaleString('es-MX');
            }
            if (filesEl) {
                filesEl.textContent = (data.total_files || 0).toLocaleString('es-MX');
            }
        } catch (err) {
            console.warn('No se pudo cargar el contador global', err);
        }
    }

    async function handleFile(file) {
        if (isProcessing) return;

        if (!file.name.toLowerCase().endsWith('.pdf')) {
            errorMessage.textContent = 'Por favor selecciona un archivo PDF válido (.pdf).';
            showState(STATES.ERROR);
            return;
        }

        if (file.size > MAX_FILE_SIZE_BYTES) {
            errorMessage.textContent = 'El archivo supera el límite máximo permitido de 50 MB.';
            showState(STATES.ERROR);
            return;
        }

        fileName.textContent = file.name;
        fileSize.textContent = formatFileSize(file.size);
        fileInfo.classList.remove('hidden');

        const formData = new FormData();
        formData.append('file', file);

        try {
            showState(STATES.UPLOADING);
            await sleep(250);

            showState(STATES.READING);
            await sleep(200);

            showState(STATES.PROCESSING);

            const response = await fetch(uploadEndpoint, {
                method: 'POST',
                body: formData,
            });

            if (!response.ok) {
                let message = `Error en el servidor (HTTP ${response.status})`;
                try {
                    const result = await response.json();
                    if (result?.message) message = result.message;
                } catch {
                    // Respuesta no-JSON (ej. proxy o túnel)
                }
                throw new Error(message);
            }

            showState(STATES.EXPORTING);

            const extractedCount = parseInt(response.headers.get('X-Records-Count') || '0', 10);
            const contentDisposition = response.headers.get('Content-Disposition');
            if (contentDisposition) {
                const match = contentDisposition.match(/filename\*?=(?:UTF-8''|"?)([^";]+)"?/i);
                if (match && match[1]) {
                    currentFileName = decodeURIComponent(match[1]);
                }
            }

            currentBlob = await response.blob();
            await sleep(200);

            if (recordCount) {
                const label = extractedCount === 1 ? 'registro extraído' : 'registros extraídos';
                recordCount.textContent = `${extractedCount.toLocaleString('es-MX')} ${label}`;
            }
            if (outputFileNameEl) {
                outputFileNameEl.textContent = currentFileName;
            }

            showState(STATES.SUCCESS);
            loadCounter();
        } catch (error) {
            console.error('Error al procesar PDF:', error);
            errorMessage.textContent = error.message || 'Ocurrió un error al procesar el archivo.';
            showState(STATES.ERROR);
        }
    }

    // Eventos de Drag & Drop
    uploadZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        if (!isProcessing) uploadZone.classList.add('drag-over');
    });

    uploadZone.addEventListener('dragleave', (e) => {
        e.preventDefault();
        uploadZone.classList.remove('drag-over');
    });

    uploadZone.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadZone.classList.remove('drag-over');
        if (isProcessing) return;
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            handleFile(files[0]);
        }
    });

    // Click en zona de carga solo cuando está en estado IDLE
    uploadZone.addEventListener('click', (e) => {
        if (isProcessing) return;
        if (e.target === btnSelect || uploadContent.contains(e.target)) {
            fileInput.click();
        }
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFile(e.target.files[0]);
        }
    });

    btnReset.addEventListener('click', (e) => {
        e.stopPropagation();
        resetUI();
    });

    btnRetry.addEventListener('click', (e) => {
        e.stopPropagation();
        resetUI();
    });

    btnDownload.addEventListener('click', (e) => {
        e.stopPropagation();
        triggerDownload();
    });

    showState(STATES.IDLE);
    loadCounter();
});
