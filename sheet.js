function getQueryParam(param) {
    const urlParams = new URLSearchParams(window.location.search);
    return urlParams.get(param);
}

const sheetName = getQueryParam('sheetName');
const fileUrl = getQueryParam('fileUrl');

(async () => {
    if (!fileUrl || !sheetName) {
        alert("Invalid sheet data.");
        return;
    }

    try {
        const response = await fetch(fileUrl);
        const arrayBuffer = await response.arrayBuffer();

        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
        const sheet = workbook.Sheets[sheetName];

        if (!sheet) {
            alert("Sheet not found.");
            return;
        }

        const html = XLSX.utils.sheet_to_html(sheet);
        const sheetContentDiv = document.getElementById('sheet-content');
        sheetContentDiv.innerHTML = html;
    } catch (error) {
        console.error("Error loading Excel file:", error);
        alert("Failed to load the Excel sheet. Please check the URL and try again.");
    }
})();
