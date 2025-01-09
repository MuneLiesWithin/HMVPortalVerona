const loadElement = document.getElementById("loadanime")

async function uploadExcel() {
    const inputFile = document.getElementById("fileUpload").files[0];

    if (!inputFile) {
        flashMessage("warning", "Por favor insira um arquivo");
        return;
    }

    // Path to the reference Excel file
    const referenceFilePath = "Excel/Excel.xlsx";

    // Loading gambiarra
    loadElement.classList.add("lds-ring")

    try {
        // Read and parse the uploaded Excel file
        const uploadedFileColumns = await getExcelColumns(inputFile);

        // Fetch and parse the reference Excel file
        const referenceFileResponse = await fetch(referenceFilePath);
        if (!referenceFileResponse.ok) {
            console.log("Error fetching reference Excel file.");
            return;
        }
        const referenceFileBlob = await referenceFileResponse.blob();
        const referenceFileColumns = await getExcelColumns(referenceFileBlob);

        // Compare columns
        if (arraysEqual(uploadedFileColumns, referenceFileColumns)) {
            // Save the uploaded file as Excel2.xlsx
            saveUploadedFile(inputFile, "Excel/Excel2.xlsx");
            flashMessage("success", "Arquivo subido com sucesso")
            loadElement.classList.remove("lds-ring")
        } else {
            console.log("An error occurred: Column names do not match.");
            flashMessage("warning", "Arquivo incompatível")
            loadElement.classList.remove("lds-ring")
        }
    } catch (error) {
        console.error("An error occurred:", error);
        flashMessage("error", "Ocorreu um erro...")
        loadElement.classList.remove("lds-ring")
    }
}

function getExcelColumns(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = function (e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array" });
            const firstSheetName = workbook.SheetNames[0];
            const firstSheet = workbook.Sheets[firstSheetName];
            const columns = XLSX.utils.sheet_to_json(firstSheet, { header: 1 })[0];

            resolve(columns);
        };

        reader.onerror = function (e) {
            reject(e);
        };

        reader.readAsArrayBuffer(file);
    });
}

function arraysEqual(arr1, arr2) {
    if (arr1.length !== arr2.length) return false;
    return arr1.every((value, index) => value === arr2[index]);
}

async function saveUploadedFile(file) {
    const formData = new FormData();
    formData.append("file", file);

    try {
        const response = await fetch("https://localhost:7232/api/excel/upload", {
            method: "POST",
            body: formData,
        });

        const result = await response.text();
        console.log(result);
    } catch (error) {
        console.error("An error occurred during file upload:", error);
    }
}


function flashMessage(type, message) {
    const messageElement = document.getElementById("message");
    messageElement.classList.add(type);
    messageElement.innerHTML = message;
    setTimeout(() => {
        messageElement.innerHTML = "";
        messageElement.classList.remove(type);
    }, 3000);
}