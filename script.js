const dropdown = document.getElementById("myDropdown");
const selected = dropdown.querySelector(".dropdown-selected");

const statusText = document.getElementById("status");

const startBtn = document.getElementById("start-btn");
const endBtn = document.getElementById("end-btn");
const refreshBtn = document.getElementById("refresh-btn");

let adapterName = null;

// Toggle dropdown
selected.addEventListener("click", () => {
  dropdown.classList.toggle("open");
});

// Close if clicked outside
document.addEventListener("click", (e) => {
  if (!dropdown.contains(e.target)) {
    dropdown.classList.remove("open");
  }
});

function updateDropdownOptions(array) {
    const options = dropdown.querySelector(".dropdown-options");

    options.innerHTML = "";

    array.forEach(option => {
        const optionDiv = document.createElement("div");
        optionDiv.classList.add("dropdown-option");

        optionDiv.textContent = option;
        optionDiv.setAttribute("data-value", option);

        options.appendChild(optionDiv);
    });

    const optionItems = dropdown.querySelectorAll(".dropdown-option");

    optionItems.forEach(option => {
        option.addEventListener("click", () => {
            selected.textContent = option.textContent;
            adapterName = option.dataset.value;
            dropdown.classList.remove("open");

            pywebview.api.save_adapter(adapterName).then(response => {
                console.log(response);
            })
        });
    });

}

window.addEventListener("pywebviewready", () => {
    startBtn.addEventListener("click", () => {
        statusText.textContent = `Hotspot running on ${adapterName}.`;
        
        pywebview.api.start_hotspot(adapterName).then(response => {
            console.log(response);
        })
    })
});

window.addEventListener("pywebviewready", () => {
    endBtn.addEventListener("click", () => {
        statusText.textContent = "Hotspot closed.";
    })
});

window.addEventListener("pywebviewready", () => {
    refreshBtn.addEventListener("click", () => {
        statusText.textContent = "Refreshing...";

        pywebview.api.refresh().then(response => {
            updateDropdownOptions(response[0]);
            statusText.textContent = "Loaded.";
        });
    })
});

window.addEventListener("pywebviewready", () => {
    pywebview.api.refresh().then(response => {
        updateDropdownOptions(response[0]);
        statusText.textContent = `Loaded.`;
        selected.textContent = response[1];
        adapterName = response[1];
        dropdown.classList.remove("open");

    });
});