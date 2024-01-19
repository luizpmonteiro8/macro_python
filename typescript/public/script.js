async function handleFile() {
  const processButton = document.getElementById("processButton");

  const input = document.getElementById("excelFile");
  const file = input.files[0];

  const configSelect = document.getElementById("configSelect");
  const selectedConfigFile = configSelect.value;

  if (file) {
    const formData = new FormData();
    formData.append("excelFile", file);

    const loadingDiv = document.getElementById("loading");
    const errorDiv = document.getElementById("error");

    // Oculta a div de erro antes de iniciar a requisição
    errorDiv.style.display = "none";

    // Mostra o indicador de carregamento
    loadingDiv.style.display = "block";
    processButton.disabled = true;

    try {
      const response = await axios.post(
        "http://127.0.0.1:3000/process-excel",
        formData,
        {
          headers: {
            configFile: selectedConfigFile,
          },
        }
      );

      const data = response.data;

      // Oculta o indicador de carregamento
      loadingDiv.style.display = "none";
      processButton.disabled = false;

      if (data.success) {
        // Faça algo com os dados retornados
      } else {
        // Exibe a mensagem de erro na div "error"
        errorDiv.innerHTML = `<div class="mt-3 alert alert-danger">${data.error}</div>`;
        // Exibe a div de erro
        errorDiv.style.display = "block";
      }
    } catch (error) {
      console.log(error);
      console.error("Erro ao processar o arquivo Excel:", error.message);

      // Exibe a mensagem de erro na div "error"
      errorDiv.innerHTML = `<div class="mt-3 alert alert-danger">${error.message}</div>`;
      // Exibe a div de erro
      errorDiv.style.display = "block";

      // Oculta o indicador de carregamento em caso de erro
      loadingDiv.style.display = "none";
      processButton.disabled = false;
    }
  } else {
    alert("Selecione um arquivo Excel.");
  }
}

function getConfigFiles() {
  fetch("http://127.0.0.1:3000/config-files")
    .then((response) => response.json())
    .then((data) => {
      const configSelect = document.getElementById("configSelect");

      if (data.success) {
        // Limpa as opções existentes
        configSelect.innerHTML = "";

        // Adiciona as novas opções do arquivo de configuração
        data.files.forEach((file) => {
          const option = document.createElement("option");
          option.text = file;
          configSelect.add(option);
        });
      } else {
        // Exibe uma mensagem de erro se a solicitação falhar
        console.error(
          "Erro ao obter a lista de arquivos de configuração:",
          data.error
        );
      }
    })
    .catch((error) => {
      console.error(
        "Erro ao obter a lista de arquivos de configuração:",
        error
      );
    });
}
function updateFileName() {
  const input = document.getElementById("excelFile");
  const label = document.querySelector(".custom-file-label");
  const fileName = input.files[0].name;
  label.textContent = fileName;
}

function editConfig() {
  const configSelect = document.getElementById("configSelect");
  const selectedConfigFile = configSelect.value;

  if (selectedConfigFile) {
    // Redireciona para a página de edição com o nome do arquivo de configuração
    window.location.href = `/edit-config?file=${selectedConfigFile}`;
  } else {
    alert("Selecione um arquivo de configuração para editar.");
  }
}
