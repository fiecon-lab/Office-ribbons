/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// Initialize an empty array to store captured ranges

let capturedRanges = [];

// document.addEventListener('DOMContentLoaded', () => {
//     // Your code here
//     const editableDiv = document.getElementById('your-editable-div-id');

//     editableDiv.addEventListener('keydown', handle_enter);

//   });

Office.onReady((info) => {
  // Office.actions.associate("PowerupsHotkey", () => {
  //   callHotkey();
  // });

  // Display the shortcut conflict dialog for testing.
  Office.actions.associate("TestConflict", () => {
    console.log("Display the shortcut conflict dialog for testing.");
  });

  // Configure the keyboard shortcut to open the task pane.
  Office.actions.associate("ShowTaskpane", () => {
    // return Office.addin
    //   .showAsTaskpane()
    //   .then(() => {
    //     return;
    //   })
    //   .catch((error) => {
    //     return error.code;
    //   });
    console.log("HI!");
  });

  // Configure the keyboard shortcut to close the task pane.
  Office.actions.associate("HideTaskpane", () => {
    return Office.addin
      .hide()
      .then(() => {
        return;
      })
      .catch((error) => {
        return error.code;
      });
  });

  // Configure the keyboard shortcut to run an action that's specific to the current Office host.
  Office.actions.associate("RunAction", () => {
    if (info.host === Office.HostType.Excel) {
      console.log("Excel add-in loaded and ready.");
    }
  });

  if (info.host === Office.HostType.Excel) {
    console.log("Excel add-in loaded and ready.");
  }
});

// ERROR HANDLING

// Global error handler - doesn't work
// window.onerror = function (message, source, lineno, colno, error) {
//   console.error("Uncaught error:", error);
//   showErrorMessage("Uncaught error:" + error);
//   // Log error or display to user
//   return true; // Prevents the firing of the default event handler
// };

const tryCatch = async (callback) => {
  try {
    // await required otherwise throws global error
    const result = await callback();
    return [result, null];
  } catch (error) {
    console.log(error);
    showErrorMessage(error);
    return [null, error];
  }
};

// Cria button

const openCria = () =>
  tryCatch(async () => {
    window.open("https://cria.fiecon.com/", "_blank");
  });

const captureAddress = async () =>
  tryCatch(async () => {
    // Your async function logic here

    if (capturedRanges.length >= 5) {
      showErrorMessage(
        "Unable to save more than 5 addresses at once. Unload the current addresses to the change log before continuing."
      );
      return;
    }

    await Excel.run(async (context) => {
      // get selected range and current worksheet
      const range = context.workbook.getSelectedRange();

      // load and sync required data
      range.load("address, worksheet, worksheet/name");
      await context.sync();

      const ws = range.worksheet;

      // construct data
      const sheet = ws.name;
      const address = range.address.split("!").pop();
      const fullAddress = range.address;
      const description = "Description of change..."; // description placeholder

      const capturedData = { sheet, address, fullAddress, description, inserted: false };
      capturedRanges.push(capturedData);

      updateCardContainer(true);
    });
  });

const updateCardContainer = async (focus_first = false) =>
  tryCatch(async () => {
    const cardContainer = document.getElementById("cardContainer");
    cardContainer.innerHTML = ""; // Clear existing cards

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

        // focus on the first input if focus_first selected
        if (focus_first && index === 0) {
          setTimeout(() => {
            cardInput.focus();
            window.getSelection().selectAllChildren(cardInput);
          }, 0);
        }

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
  });

// Function to delete a card from the capturedRanges array and update the card container
const deleteCard = async (index) =>
  tryCatch(async () => {
    // Remove the card at the specified index from the capturedRanges array
    capturedRanges.splice(index, 1);
    // Update the card container to reflect the changes in the capturedRanges array
    updateCardContainer();
  });

const insertSingleCard = async (index) =>
  tryCatch(async () => {
    if (capturedRanges[index].inserted) {
      // If already inserted, clicking button resets (instead of inserting again)
      capturedRanges[index].inserted = false;
      updateCardContainer();
    } else {
      insertAddress(index);
    }
  });

const insertAllCards = async () =>
  tryCatch(async () => {
    let i;
    for (i = 0; i < capturedRanges.length; i++) {
      capturedRanges[i].inserted = true;
    }

    showErrorMessage("Inserting all!");
    updateCardContainer();
  });

const deleteAllCards = async () =>
  tryCatch(async () => {
    capturedRanges = [];
    updateCardContainer();
  });

const showTab = async (tabIndex) =>
  tryCatch(async () => {
    const tabs = ["home-tab", "address-clipper", "suggestions-tab"];

    // Hide all tabs
    tabs.forEach((id) => {
      document.getElementById(id).classList.remove("active");
    });

    // Show the selected tab
    document.getElementById(tabs[tabIndex - 1]).classList.add("active");
  });

function showErrorMessage(message) {
  const popup = document.createElement("div");
  popup.classList.add("error-popup");

  const messageElement = document.createElement("span");
  messageElement.textContent = message;

  const closeButton = document.createElement("button");
  closeButton.innerHTML = "&times;"; // Unicode character for "X"
  closeButton.addEventListener("click", () => {
    popup.classList.add("fade-out");
    setTimeout(() => {
      popup.remove();
    }, 500);
  });

  popup.appendChild(messageElement);
  popup.appendChild(closeButton);

  const container = document.querySelector(".error-popup-container");
  container.prepend(popup);

  setTimeout(() => {
    popup.classList.add("fade-out");
    setTimeout(() => {
      popup.remove();
    }, 500);
  }, 5000);
}

const insertAddress = async (index) =>
  tryCatch(async () => {
    await Excel.run(async (context) => {
      const activeCell = context.workbook.getActiveCell();

      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      activeSheet.load("name"); // Explicitly load the 'name' property
      await context.sync(); // Sync to load the properties

      // Execute only if the active sheet is "Change log"
      if (activeSheet.name.toLowerCase() !== "change log") {
        console.log("Action not allowed. The active sheet is not 'Change log'.");
        // throw new Error("Unable to insert address. Destination sheet must be 'Change log'.");
        throw Error("Unable to insert address. Destination sheet must be 'Change log'.");
      }

      const data = capturedRanges[index];
      const initials = document.getElementById("initialsInput").value || "N/A";

      // Create a horizontal array with initials and address
      const valueArray = [
        [data.sheet, getAddressLinkDynamic(data.sheet, data.address, data.fullAddress), data.description, initials],
      ];

      console.log(valueArray);

      // Get a range that's two cells wide starting from the active cell
      const targetRange = activeCell.getResizedRange(0, 3);

      targetRange.load("valueTypes");

      await context.sync();

      let isEmpty = targetRange.valueTypes.every((row) =>
        row.every((cell) => cell === Excel.RangeValueType.empty || cell === null || cell === "")
      );

      if (!isEmpty) {
        targetRange.select();
        // throw new Error("Unable to insert address. Destination must be empty.");
        throw Error("Unable to insert address. Destination must be empty.");
      }

      // Set the values of the range
      targetRange.values = valueArray;

      // Success!
      const nextRowCell = targetRange.getCell(0, 0).getOffsetRange(1, 0);
      nextRowCell.select();

      await context.sync();

      capturedRanges[index].inserted = true;
    });
  });

function getAddressLinkDynamic(sht, addr, fullAddr) {
  // Construct the LET formula with HYPERLINK
  const linkFormula = `= LET(rng, ${fullAddr}, sht, TEXTAFTER(CELL("filename", rng), "]"), addr, IF(ROWS(rng) + COLUMNS(rng)=2, ADDRESS(ROW(rng), COLUMN(rng)), ADDRESS(MIN(ROW(rng)), MIN(COLUMN(rng))) & ":" & ADDRESS(MAX(ROW(rng)), MAX(COLUMN(rng)))), dynamic_link, HYPERLINK("#'" & sht & "'!" & addr, "↗️" & SUBSTITUTE(addr, "$", "")), IFERROR(dynamic_link, HYPERLINK("#'${sht}'!${addr}","[static!] ↗️${addr.replace(
    "$",
    ""
  )}")))`;

  return linkFormula;
}

const callHotkey = () => {
  return Office.addin
    .showAsTaskpane()
    .then(() => {
      return;
    })
    .catch((error) => {
      return error.code;
    });
};
