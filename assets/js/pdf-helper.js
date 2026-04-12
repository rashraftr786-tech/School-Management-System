const PDFHelper = {
    generate: async (data, isSingle) => {
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF();
        // ... (PDF logic using jsPDF and html2canvas)
        console.log("PDF generated for", isSingle ? "Student" : "Full Table");
    }
};
