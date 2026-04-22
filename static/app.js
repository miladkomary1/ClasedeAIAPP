document.addEventListener('DOMContentLoaded', () => {
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const uploadContent = document.querySelector('.upload-content');
    const processingState = document.getElementById('processing-state');
    const resultMessage = document.getElementById('result-message');

    // Handle click on drop zone to open file dialog
    dropZone.addEventListener('click', () => {
        fileInput.click();
    });

    // Handle drag events
    ['dragover', 'dragenter'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropZone.classList.add('dragover');
        }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropZone.classList.remove('dragover');
        }, false);
    });

    // Handle drop event
    dropZone.addEventListener('drop', (e) => {
        const files = e.dataTransfer.files;
        handleFiles(files);
    });

    // Handle file selection from input
    fileInput.addEventListener('change', function() {
        handleFiles(this.files);
    });

    function handleFiles(files) {
        if (files.length === 0) return;
        
        const file = files[0];
        
        // Validate file extension
        if (!file.name.toLowerCase().endsWith('.dxf')) {
            showMessage('Error: Por favor sube solo archivos .dxf', 'error');
            return;
        }

        uploadFile(file);
    }

    function showMessage(msg, type) {
        resultMessage.textContent = msg;
        resultMessage.className = `result-message show ${type}`;
        
        // Hide message after 5 seconds if it's an error
        if(type === 'error') {
            setTimeout(() => {
                resultMessage.classList.remove('show');
            }, 5000);
        }
    }

    function toggleProcessingState(isProcessing) {
        if (isProcessing) {
            uploadContent.style.display = 'none';
            processingState.style.display = 'flex';
            dropZone.style.pointerEvents = 'none';
            resultMessage.classList.remove('show');
        } else {
            uploadContent.style.display = 'block';
            processingState.style.display = 'none';
            dropZone.style.pointerEvents = 'auto';
            fileInput.value = ''; // Reset input
        }
    }

    async function uploadFile(file) {
        const formData = new FormData();
        formData.append('file', file);

        toggleProcessingState(true);

        try {
            const response = await fetch('/api/upload', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                let errorText = 'Error al procesar el archivo.';
                try {
                    const error = await response.json();
                    errorText = error.error || errorText;
                } catch(e) {}
                
                throw new Error(errorText);
            }

            // Handle successful response (json expected)
            const result = await response.json();
            
            if (result.success) {
                // Populate rows
                const tbody = document.getElementById('results-body');
                tbody.innerHTML = '';
                
                result.rows.forEach(row => {
                    const tr = document.createElement('tr');
                    // Add subtle warning style if there are flags (like range)
                    if (row.flags && row.flags.length > 0) {
                        tr.style.background = 'rgba(232, 162, 2, 0.15)'; // pale amber
                    }
                    tr.innerHTML = `
                        <td><strong>${row.mark || '-'}</strong></td>
                        <td>${row.note}</td>
                        <td>${row.total_bars}</td>
                        <td>Ø${row.dia_mm}</td>
                        <td>${row.total_length_m.toFixed(2)}</td>
                        <td>${row.total_weight_kg.toFixed(2)}</td>
                    `;
                    tbody.appendChild(tr);
                });
                
                // Show results container, hide drop container
                document.getElementById('drop-zone').style.display = 'none';
                document.getElementById('results-container').style.display = 'block';
                
                // Setup download button
                document.getElementById('download-btn').onclick = () => {
                   window.location.href = result.download_url;
                };
                
                // Setup new upload button
                document.getElementById('new-upload-btn').onclick = () => {
                   document.getElementById('results-container').style.display = 'none';
                   document.getElementById('drop-zone').style.display = 'block';
                   document.getElementById('file-input').value = '';
                };
                
                showMessage('¡Archivo analizado con éxito!', 'success');
            } else {
                throw new Error(result.error || 'Error desconocido.');
            }
            
            // Hide success message after 4 seconds
            setTimeout(() => {
                resultMessage.classList.remove('show');
            }, 4000);
            
        } catch (error) {
            console.error('Upload Error:', error);
            showMessage(error.message, 'error');
        } finally {
            toggleProcessingState(false);
        }
    }
});
