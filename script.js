// Function to verify everything is linked
console.log("Result Management System: Script loaded.");

// 1. Wrap in DOMContentLoaded to ensure HTML elements exist
document.addEventListener('DOMContentLoaded', () => {
    console.log("DOM fully loaded.");
});

// --- PDF GENERATION LOGIC ---

function generateReport(index) {
    // If index is undefined, it means the main "Download PDF" button was clicked
    // If index is a number, it's a specific student row
    console.log("Generating PDF for index:", index);

    // Initialize jsPDF (Standard for v2.5.1)
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();

    if (index !== undefined) {
        // Individual Student Report
        const table = document.getElementById('resultsTable');
        const row = table.rows[index + 1]; // +1 to skip header
        const name = row.cells[1].innerText;
        
        doc.setFontSize(18);
        doc.text("Student Result Report", 20, 20);
document.addEventListener('DOMContentLoaded', () => {
    console.log("Result Management System: Ready.");
});

// --- PDF GENERATION LOGIC ---
async function generateReport(index) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('p', 'mm', 'a4');

    if (index !== undefined && index !== null) {
        // Individual Student Report
        const row = document.getElementById('tableBody').rows[index];
        const name = row.cells[1].innerText;
        
        doc.setFontSize(18);
        doc.text("Student Result Report", 20, 20);
        doc.setFontSize(12);
        doc.text(`Name: ${name}`, 20, 40);
        doc.text(`Father's Name: ${row.cells[2].innerText}`, 20, 50);
        doc.text(`Roll No: ${row.cells[3].innerText}`, 20, 60);
        doc.text(`Grade: ${row.cells[6].innerText}`, 20, 70);
        doc.save(`Result_${name}.pdf`);
    } else {
        // Full Table Export (High Quality)
        const table = document.getElementById("resultsTable");
        const canvas = await html2canvas(table, { scale: 2 });
        const imgData = canvas.toDataURL('image/png');
        
        const pdfWidth = doc.internal.pageSize.getWidth();
        const pdfHeight = (canvas.height * pdfWidth) / canvas.width;
        
        doc.addImage(imgData, 'PNG', 0, 10, pdfWidth, pdfHeight);
        doc.save("School_Results_Full.pdf");
    }
}

// --- PHOTO UPLOAD LOGIC ---
let selectedStudentIndex = null;

function openPhotoModal(index) {
    selectedStudentIndex = index;
    document.getElementById('photoModal').style.display = 'block';
}

function closeModal() {
    document.getElementById('photoModal').style.display = 'none';
}

function performUpload() {
    const fileInput = document.getElementById('photoFile');
    const file = fileInput.files[0];

    if (!file || selectedStudentIndex === null) {
        alert("Please select a photo first.");
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        const row = document.getElementById('tableBody').rows[selectedStudentIndex];
        const imgTag = row.cells[0].querySelector('img');
        imgTag.src = e.target.result;
        alert("Photo updated successfully!");
        closeModal();
    };
    reader.readAsDataURL(file);
}

// --- EXCEL LOGIC ---
function processExcel() {
    const file = document.getElementById('excelFile').files[0];
    if (!file) return alert("Please select an Excel file.");

    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
        renderTable(jsonData);
    };
    reader.readAsArrayBuffer(file);
}

function renderTable(data) {
    const tableBody = document.getElementById('tableBody');
    tableBody.innerHTML = data.map((student, index) => `
        <tr>
            <td><img src="https://via.placeholder.com/40" style="border-radius:50%; width:40px; height:40px;"></td>
            <td>${student['Student Name'] || 'N/A'}</td>
            <td>${student["Father's Name"] || 'N/A'}</td>
            <td>${student['Roll No'] || 'N/A'}</td>
            <td>${student['Class'] || 'N/A'}</td>
            <td>${student['Total (%)'] || '0'}%</td>
            <td>${student['Grade'] || 'N/A'}</td>
            <td>${student['Rank'] || 'N/A'}</td>
            <td>
                <button onclick="openPhotoModal(${index})" class="btn-sm">Photo</button>
                <button onclick="generateReport(${index})" class="btn-sm">PDF</button>
            </td>
        </tr>
    `).join('');
}
        
