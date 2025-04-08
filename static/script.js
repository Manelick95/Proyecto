document.addEventListener("DOMContentLoaded", function () {
    const form = document.getElementById("uploadForm");
    const mensaje = document.getElementById("mensaje");

    form.addEventListener("submit", function (event) {
        event.preventDefault();
        const formData = new FormData(form);

        fetch("/upload", {
            method: "POST",
            body: formData
        })
        .then(response => {
            if (!response.ok) {
                throw new Error("Error en la respuesta del servidor");
            }
            return response.blob();
        })
        .then(blob => {
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.style.display = "none";
            a.href = url;
            a.download = "archivo_modificado.xls";  // 🔧 Cambiado a .xls
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            mensaje.innerHTML = "✅ Archivo descargado correctamente.";
        })
        .catch(error => {
            console.error("Error:", error);
            mensaje.innerHTML = "❌ Hubo un error al procesar el archivo.";
        });
    });
});
