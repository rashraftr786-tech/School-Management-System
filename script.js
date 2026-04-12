// Function to verify everything is linked
console.log("Result Management System: Script loaded.");

// 1. Global variables to hold student data and current selection
let studentData = [];
let currentStudentIndex = null;

// 2. The function called by the "Process File" button
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
            
            // Store data globally so other functions can access it
            studentData = XLSX.utils.sheet_to_json(firstSheet);
            
            if (studentData.length > 0) {
                renderTable(studentData);
                alert("Upload successful! Found " + studentData.length + " students.");
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

// 3. Render Table with correct function calls
function renderTable(data) {
    const tableBody = document.getElementById('tableBody');
    tableBody.innerHTML = ''; 

    data.forEach((student, index) => {
        const row = `
            <tr>
                <td><img id="img-${index}" src="https://via.placeholder.com/40" alt="Student" style="border-radius: 50%;"></td>
                <td>${student['Student Name'] || 'N/A'}</td>
                <td>${student["Father's Name"] || 'N/A'}</td>
                <td>${student['Roll No'] || 'N/A'}</td>
                <td>${student['Class'] || 'N/A'}</td>
                <td>${student['Total (%)'] || '0'}%</td>
                <td>${student['Grade'] || 'N/A'}</td>
                <td>${student['Rank'] || 'N/A'}</td>
                <td>
                    <button onclick="openPhotoModal(${index})" class="btn-sm">Photo</button>
                    <button onclick="downloadPDF(${index})" class="btn-sm">PDF</button>
                </td>
            </tr>
        `;
        tableBody.innerHTML += row;
    });
}

// 4. FIXED: PDF Download Logic
async function downloadPDF(index) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();
    
    // If an index is passed, download individual result, else download the whole table
    const targetElement = (index !== undefined) ? 
        document.querySelectorAll('#tableBody tr')[index] : 
        document.getElementById('resultsTable');

    if (!targetElement) {
        alert("No data available to export.");
        return;
    }

    try {
        const canvas = await html2canvas(targetElement);
        const imgData = canvas.toDataURL('image/png');
        const imgProps = doc.getImageProperties(imgData);
        const pdfWidth = doc.internal.pageSize.getWidth();
        const pdfHeight = (imgProps.height * pdfWidth) / imgProps.width;

        doc.text("GPS Pethgam Wagoora - Result Report", 10, 10);
        doc.addImage(imgData, 'PNG', 0, 20, pdfWidth, pdfHeight);
        doc.save(`Result_${index !== undefined ? studentData[index]['Student Name'] : 'School'}.pdf`);
    } catch (error) {
        console.error("PDF Generation Error:", error);
        alert("Error generating PDF. Check console.");
    }
}

// 5. FIXED: Photo Upload Compatibility
function openPhotoModal(index) {
    currentStudentIndex = index;
    document.getElementById('photoModal').style.display = 'block';
}

function uploadPhoto() {
    const fileInput = document.getElementById('photoFile');
    const file = fileInput.files[0];
