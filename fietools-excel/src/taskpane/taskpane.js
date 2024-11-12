// Initialize an empty array to store captured ranges

let capturedRanges = [];

// document.addEventListener('DOMContentLoaded', () => {
//     // Your code here
//     const editableDiv = document.getElementById('your-editable-div-id');

//     editableDiv.addEventListener('keydown', handle_enter);

//   });

function capture_address() {
  let i = capturedRanges.length + 1;
  let sheet = `Sheet${i}`;
  let address = `A${i}:B${i * 10}`;
  let fullAddress = `'${sheet}'!${address}`;
  let description = `Description of change to ${fullAddress}`;

  const capturedData = { sheet, address, fullAddress, description, inserted: false };
  capturedRanges.push(capturedData);

  updateCardContainer();
}

function updateCardContainer() {
  const cardContainer = document.getElementById("cardContainer");
  cardContainer.innerHTML = ""; // Clear existing cards

  let firstInput = null;

  console.log(capturedRanges.length);

  if (capturedRanges.length == 0) {
    cardContainer.innerHTML = `<p id="no-addresses-message">Click "Capture Address" to get started.</p>`;
    return;
  }

  capturedRanges
    .slice()
    .reverse()
    .forEach((data, index) => {
      // Card root div
      const cardInstance = document.createElement("div");
      cardInstance.className = "card-instance";

      // Add insert button
      const insertBtn = document.createElement("button");
      insertBtn.className = "card-insert" + (data.inserted ? " complete" : "");
      insertBtn.innerHTML = data.inserted
        ? `
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#4CAF50" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
                <circle cx="12" cy="12" r="11"/>
                <path d="M8 12l3 3 5-5"/>
            </svg>`
        : `
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="black" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <polyline points="15 18 9 12 15 6"></polyline>
            </svg>`;
      insertBtn.onclick = () => insertSingleCard(capturedRanges.length - 1 - index);
      cardInstance.appendChild(insertBtn);

      // Card root div
      const card = document.createElement("div");
      card.className = "card";

      // Create card header
      const cardHeader = document.createElement("div");
      cardHeader.className = "card-header";

      // Create address wrapper
      const addressWrapper = document.createElement("div");
      addressWrapper.className = "card-address-wrapper";

      // Create sheet button
      const sheetButton = document.createElement("button");
      sheetButton.className = "card-address";
      sheetButton.textContent = data.sheet;

      // Create range button
      const rangeButton = document.createElement("button");
      rangeButton.className = "card-address";
      rangeButton.textContent = data.address;

      // Add buttons to address wrapper
      addressWrapper.appendChild(sheetButton);
      addressWrapper.appendChild(rangeButton);

      // Create delete button
      const deleteButton = document.createElement("button");
      deleteButton.className = "card-delete-button";
      deleteButton.setAttribute("aria-label", "Delete");
      deleteButton.setAttribute("type", "button");
      deleteButton.innerHTML = `<img src="../../assets/icon_delete_svg.svg" alt="Delete icon">`;
      deleteButton.onclick = () => deleteCard(capturedRanges.length - 1 - index);

      // Add elements to card header
      cardHeader.appendChild(addressWrapper);
      cardHeader.appendChild(deleteButton);

      // Create card input
      const cardInput = document.createElement("div");
      cardInput.className = "card-input";
      cardInput.textContent = data.description;
      cardInput.setAttribute("contenteditable", "true");
      cardInput.setAttribute("placeholder", "Enter description...");
      cardInput.addEventListener("input", function () {
        capturedRanges[capturedRanges.length - 1 - index].description = this.textContent;
      });
      cardInput.addEventListener("keydown", function (e) {
        if (e.key === "Enter" && !e.shiftKey) {
          this.blur();
          window.getSelection().removeAllRanges();
        }
      });

      // Create card footer
      const cardFooter = document.createElement("div");
      cardFooter.className = "card-footer";
      cardFooter.innerHTML = `<p style="font-size:0.75rem">Use <span style="background-color: antiquewhite; border-radius: 2px; padding:0.1rem 0.2rem">shift+return</span> for a new line</p>`;
      cardFooter.style.display = "none"; // Initially hidden
      // Add event listeners to show/hide cardFooter based on cardInput focus
      cardInput.addEventListener("focus", function () {
        cardFooter.style.display = "block";
      });
      cardInput.addEventListener("blur", function () {
        cardFooter.style.display = "none";
      });
      
      // Add all elements to main card
      card.appendChild(cardHeader);
      card.appendChild(cardInput);
      card.appendChild(cardFooter);

      cardInstance.append(card);
      cardContainer.appendChild(cardInstance);
    });

  // Focus and select text of the first (most recent) input
  //   if (firstInput) {
  //     setTimeout(() => {
  //       firstInput.focus();
  //       firstInput.select();
  //     }, 0);
  //   }
}

// Function to delete a card from the capturedRanges array and update the card container
function deleteCard(index) {
  // Remove the card at the specified index from the capturedRanges array
  capturedRanges.splice(index, 1);
  // Update the card container to reflect the changes in the capturedRanges array
  updateCardContainer();
}

function insertSingleCard(index) {
  capturedRanges[index].inserted = !capturedRanges[index].inserted;
  updateCardContainer();
}

function insertAllCards() {
  for (i = 0; i < capturedRanges.length; i++) {
    capturedRanges[i].inserted = true;
  }
  updateCardContainer();
}

function deleteAllCards() {
  capturedRanges = [];
  updateCardContainer();
}

function showTab(tabIndex) {
  const tabs = ["home-tab", "address-clipper", "suggestions-tab"];

  // Hide all tabs
  tabs.forEach((id) => {
    document.getElementById(id).classList.remove("active");
  });

  // Show the selected tab
  document.getElementById(tabs[tabIndex - 1]).classList.add("active");
}
