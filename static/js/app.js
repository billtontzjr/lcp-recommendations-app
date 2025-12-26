document.addEventListener('DOMContentLoaded', function() {
    const uploadForm = document.getElementById('uploadForm');
    const workbookDropZone = document.getElementById('workbookDropZone');
    const workbookFile = document.getElementById('workbookFile');
    const workbookFileName = document.getElementById('workbookFileName');
    const summaryDropZone = document.getElementById('summaryDropZone');
    const summaryFile = document.getElementById('summaryFile');
    const summaryFileName = document.getElementById('summaryFileName');
    const previewBtn = document.getElementById('previewBtn');
    const generateBtn = document.getElementById('generateBtn');
    const previewSection = document.getElementById('previewSection');
    const loadingSection = document.getElementById('loadingSection');
    const errorSection = document.getElementById('errorSection');
    const successSection = document.getElementById('successSection');
    const resetBtn = document.getElementById('resetBtn');

    // Drop zone handlers for workbook
    setupDropZone(workbookDropZone, workbookFile, workbookFileName, ['xlsx', 'xlsm']);
    setupDropZone(summaryDropZone, summaryFile, summaryFileName, ['docx']);

    // Click to select file
    workbookDropZone.addEventListener('click', () => workbookFile.click());
    summaryDropZone.addEventListener('click', () => summaryFile.click());

    // File input change handlers
    workbookFile.addEventListener('change', function() {
        updateFileName(this, workbookFileName, workbookDropZone);
        updateButtons();
    });

    summaryFile.addEventListener('change', function() {
        updateFileName(this, summaryFileName, summaryDropZone);
    });

    // Preview button
    previewBtn.addEventListener('click', async function() {
        if (!workbookFile.files[0]) return;

        hideAllSections();
        loadingSection.classList.remove('hidden');

        const formData = new FormData();
        formData.append('file', workbookFile.files[0]);

        try {
            const response = await fetch('/api/preview', {
                method: 'POST',
                body: formData
            });

            const data = await response.json();

            if (!response.ok) {
                throw new Error(data.error || 'Preview failed');
            }

            displayPreview(data);
            loadingSection.classList.add('hidden');
            previewSection.classList.remove('hidden');

        } catch (error) {
            showError(error.message);
        }
    });

    // Generate form submission
    uploadForm.addEventListener('submit', async function(e) {
        e.preventDefault();

        if (!workbookFile.files[0]) return;

        hideAllSections();
        loadingSection.classList.remove('hidden');

        const formData = new FormData();
        formData.append('file', workbookFile.files[0]);

        if (summaryFile.files[0]) {
            formData.append('medical_summary', summaryFile.files[0]);
        }

        try {
            const response = await fetch('/api/generate', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                const data = await response.json();
                throw new Error(data.error || 'Generation failed');
            }

            // Download the file
            const blob = await response.blob();
            const contentDisposition = response.headers.get('Content-Disposition');
            let filename = 'LCP_Recommendations.docx';

            if (contentDisposition) {
                const match = contentDisposition.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/);
                if (match && match[1]) {
                    filename = match[1].replace(/['"]/g, '');
                }
            }

            // Create download link
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);

            showSuccess(filename);

        } catch (error) {
            showError(error.message);
        }
    });

    // Reset button
    resetBtn.addEventListener('click', function() {
        uploadForm.reset();
        workbookFileName.textContent = '';
        summaryFileName.textContent = '';
        workbookDropZone.classList.remove('has-file');
        summaryDropZone.classList.remove('has-file');
        hideAllSections();
        updateButtons();
    });

    // Helper functions
    function setupDropZone(dropZone, fileInput, fileNameEl, allowedExtensions) {
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            dropZone.addEventListener(eventName, preventDefaults, false);
        });

        ['dragenter', 'dragover'].forEach(eventName => {
            dropZone.addEventListener(eventName, () => dropZone.classList.add('dragover'), false);
        });

        ['dragleave', 'drop'].forEach(eventName => {
            dropZone.addEventListener(eventName, () => dropZone.classList.remove('dragover'), false);
        });

        dropZone.addEventListener('drop', function(e) {
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                const file = files[0];
                const ext = file.name.split('.').pop().toLowerCase();

                if (allowedExtensions.includes(ext)) {
                    fileInput.files = files;
                    updateFileName(fileInput, fileNameEl, dropZone);
                    updateButtons();
                }
            }
        }, false);
    }

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    function updateFileName(input, nameEl, dropZone) {
        if (input.files[0]) {
            nameEl.textContent = input.files[0].name;
            dropZone.classList.add('has-file');
        } else {
            nameEl.textContent = '';
            dropZone.classList.remove('has-file');
        }
    }

    function updateButtons() {
        const hasFile = workbookFile.files && workbookFile.files[0];
        previewBtn.disabled = !hasFile;
        generateBtn.disabled = !hasFile;
    }

    function hideAllSections() {
        previewSection.classList.add('hidden');
        loadingSection.classList.add('hidden');
        errorSection.classList.add('hidden');
        successSection.classList.add('hidden');
    }

    function displayPreview(data) {
        const patientInfo = document.getElementById('patientInfo');
        const costSummary = document.getElementById('costSummary');

        patientInfo.innerHTML = `
            <p><strong>Patient:</strong> ${data.patient_info.patient_name || 'N/A'}</p>
            <p><strong>Date of Birth:</strong> ${data.patient_info.date_of_birth || 'N/A'}</p>
            <p><strong>Date of Injury:</strong> ${data.patient_info.date_of_injury || 'N/A'}</p>
            <p><strong>Life Expectancy:</strong> ${data.patient_info.life_expectancy || 'N/A'} years</p>
            <p><strong>Total Items:</strong> ${data.item_count}</p>
        `;

        let categoryRows = '';
        for (const [category, catData] of Object.entries(data.categories)) {
            categoryRows += `
                <tr>
                    <td>${category}</td>
                    <td>${formatCurrency(catData.annual_cost)}</td>
                    <td>${formatCurrency(catData.one_time_cost)}</td>
                </tr>
            `;
        }

        costSummary.innerHTML = `
            <table>
                <thead>
                    <tr>
                        <th>Category</th>
                        <th>Annual Cost</th>
                        <th>One-Time Cost</th>
                    </tr>
                </thead>
                <tbody>
                    ${categoryRows}
                    <tr class="totals-row">
                        <td>Total Annual</td>
                        <td colspan="2">${formatCurrency(data.totals.total_annual)}</td>
                    </tr>
                    <tr class="totals-row">
                        <td>Lifetime Annual (${data.totals.life_expectancy} years)</td>
                        <td colspan="2">${formatCurrency(data.totals.lifetime_annual)}</td>
                    </tr>
                    <tr class="totals-row">
                        <td>Total One-Time</td>
                        <td colspan="2">${formatCurrency(data.totals.total_one_time)}</td>
                    </tr>
                    <tr class="grand-total">
                        <td>GRAND TOTAL</td>
                        <td colspan="2">${formatCurrency(data.totals.grand_total)}</td>
                    </tr>
                </tbody>
            </table>
        `;
    }

    function showError(message) {
        hideAllSections();
        document.getElementById('errorMessage').textContent = message;
        errorSection.classList.remove('hidden');
    }

    function showSuccess(filename) {
        hideAllSections();
        document.getElementById('successMessage').textContent = `Document "${filename}" has been downloaded.`;
        successSection.classList.remove('hidden');
    }

    function formatCurrency(amount) {
        if (amount === null || amount === undefined) return '$0.00';
        return '$' + parseFloat(amount).toLocaleString('en-US', {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        });
    }
});
