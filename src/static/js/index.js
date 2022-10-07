var form = document.getElementById("upload-form");
var csvInput = document.getElementById("csv-input");
var modelInput = document.getElementById("model-input");

//Listening for a newly added file
csvInput.onchange = () => updateFilename(csvInput)
modelInput.onchange = () => updateFilename(modelInput)

function updateFilename(input) {
  file = new FileReader();
  //Get the corresponding paragraph
  var displayer = document.getElementById(`filename-display-${input.id.split('-')[0]}`);
  //Check if there's a file added
  try {
    file.readAsBinaryString(input.files[0]);
    //If there is a file then show the name of it
    displayer.innerText = input.files[0].name;
  } catch {
    displayer.innerText = 'Nincs fájl kiválasztva';
    return;
  }
};

//Waiting for the form to be submitted
form.addEventListener("submit", async function (event) {
  event.preventDefault();

  csv_reader = new FileReader();
  model_reader = new FileReader();


  //If theres a file added don't show any warning
  await csv_reader.readAsBinaryString(csvInput.files[0]);
  await model_reader.readAsBinaryString(modelInput.files[0]);
  document.getElementById("no-file-warn").hidden = true;
  

  var downloadLink = document.getElementById("file-downloader");
  var loader = document.getElementById("converting-loader");
  loader.hidden = false;

  var isloaded = false;
  //Waiting for the file to load
  csv_reader.onload = () => checkIfLoaded()
  model_reader.onload = () => checkIfLoaded()

  function checkIfLoaded() {
    if(isloaded) {
      sendRequest();
    }
    isloaded = true;
  }
  
  async function sendRequest() {
    //Creating the url for the request
    var url = new URL(location.origin + "/convert_to_docx");

    //Adding the binary to the form
    const formData = new FormData();
    formData.append("csv", csvInput.files[0]);
    formData.append("model", modelInput.files[0]);

    //Sending the request with fetch
    var res = await fetch(url, { method: "POST", body: formData });

    //Creating a download link for the .zip file
    const file = await res.blob();
    const fileURL = URL.createObjectURL(file);

    loader.hidden = true;

    downloadLink.setAttribute("href", fileURL);
    downloadLink.setAttribute(
      "download",
      `${csvInput.files[0].name.slice(0, -4)}.zip`
    );
    downloadLink.hidden = false;
  };
});
