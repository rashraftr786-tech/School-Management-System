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
        doc.setFontSize(12);
        doc.text(`Name: ${name}`, 20, 40);
        doc.text(`Father's Name: ${row.cells[2].innerText}`, 20, 50);
        doc.text(`Roll No: ${row.cells[3].innerText}`, 20, 60);
        doc.text(`Grade: ${row.cells[6].innerText}`, 20, 70);
        
        doc.save(`Result_${name}.pdf`);
    } else {
        // Full Table Export
        html2canvas(document.getElementById("resultsTable")).then(canvas => {
            const imgData = canvas.toDataURL('image/png');
            const imgProps = doc.getImageProperties(imgData);
            const pdfWidth = doc.internal.pageSize.getWidth();
            const pdfHeight = (imgProps.height * pdfWidth) / imgProps.width;
            
            doc.addImage(imgData, 'PNG', 0, 0, pdfWidth, pdfHeight);
            doc.save("School_Results_Full.pdf");
        });
    }
}

// --- PHOTO UPLOAD LOGIC ---

function uploadPhoto(index) {
    const fileInput = document.getElementById('photoFile');
    const file = fileInput.files[0];

    if (!file) {
        alert("Please select a photo file first.");
        return;
    }

    // Logic to update the table image preview
    const reader = new FileReader();
    reader.onload = function(e) {
        const tableBody = document.getElementById('tableBody');
        const row = tableBody.rows[index];
        const imgTag = row.cells[0].querySelector('img');
        imgTag.src = e.target.result; // Updates the placeholder with the real photo
        
        alert("Photo uploaded and updated in table!");
        if(typeof closeModal === "function") closeModal();
    };
    reader.readAsDataURL(file);
}

// --- EXISTING EXCEL LOGIC ---

function processExcel() {
    const fileInput = document.getElementById('excelFile');
    const file = fileInput.files[0];

    if (!file) {
        alert("Please select an Excel file first.");
        return;
    }

    const reader = new FileReader();
    reader.onload = function(evt) {
        try {
            const data = evt.target.result;
            const workbook = XLSX.read(data, { type: 'binary' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet);
            
            if (jsonData.length > 0) {
                renderTable(jsonData);
                alert("Upload successful! Found " + jsonData.length + " students.");
            } else {
                alert("The Excel file seems to be empty.");
            }
        } catch (err) {
            console.error("Error reading file:", err);
            alert("Could not read Excel file.");
        }
    };
    reader.readAsBinaryString(file);
}

function renderTable(data) {
    const tableBody = document.getElementById('tableBody');
    tableBody.innerHTML = ''; 

    data.forEach((student, index) => {
        const row = `
            <tr>
                <td><img src="https://via.placeholder.com/40" alt="Student" style="border-radius:50%"></td>
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
        `;
        tableBody.innerHTML += row;
    });
}

function filterTable() {
    const input = document.getElementById('searchInput').value.toUpperCase();
    const table = document.getElementById('resultsTable');
    const tr = table.getElementsByTagName('tr');

    for (let i = 1; i < tr.length; i++) {
        const td = tr[i].getElementsByTagName('td')[1]; // Student Name column
        if (td) {
            const txtValue = td.textContent || td.innerText;
            tr[i].style.display = txtValue.toUpperCase().indexOf(input) > -1 ? "" : "none";
        }
    }
}
    
