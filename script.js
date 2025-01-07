function uploadExcel() {
    console.log("TESTE...")
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

function handleKeyPress(event, buttonId) {
    if (event.key === 'Enter') {
        event.preventDefault(); 
        document.getElementById(buttonId).click(); 
    }
}