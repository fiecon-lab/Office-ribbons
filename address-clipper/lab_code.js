name: AddressClipper
description: Create a new snippet from a blank template.
host: EXCEL
api_set: {}
script:
  content: >
    $("#captureAddressBtn").on("click", () => tryCatch(run));




    async function run() {
      await Excel.run(async (context) => {
        await captureAddress(context);
      });
    }


    /** Default helper for invoking an action and handling errors. */

    async function tryCatch(callback) {
      try {
        await callback();
      } catch (error) {
        // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
        console.error(error);
      }
    }


    async function captureAddress(context: Excel.RequestContext) {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        const ws = range.worksheet;

        ws.load("name");
        range.load("address");

        await context.sync();

        const sheet = ws.name;
        const address = range.address.split("!").pop();
        const fullAddress = range.address;

        // Capture the range and additional information
        const capturedData = { sheet, address, fullAddress, description: address };
        capturedRanges.push(capturedData);

        // console.log(capturedRanges);

        // Update the UI
        updateCardContainer();
      });
    }


    function updateCardContainer() {
      const cardContainer = document.getElementById("cardContainer");
      cardContainer.innerHTML = ""; // Clear existing cards

      let firstInput = null;

      capturedRanges
        .slice()
        .reverse()
        .forEach((data, index) => {
          const card = document.createElement("div");
          card.className = "card";

          const deleteBtn = document.createElement("button");
          deleteBtn.className = "delete-btn";
          deleteBtn.innerHTML = "&times;";
          deleteBtn.onclick = () => deleteCard(capturedRanges.length - 1 - index);

          const rangeAddress = document.createElement("textarea");
          rangeAddress.className = "range-address";
          rangeAddress.value = data.description;
          rangeAddress.onchange = (e) => updateDescription(capturedRanges.length - 1 - index, e.target.value);

          if (index === 0) {
            firstInput = rangeAddress;
          }

          const menuBar = document.createElement("div");
          menuBar.className = "menu-bar";

          const infoSpan = document.createElement("span");
          infoSpan.textContent = data.fullAddress;
          // add an input

          const insertButton = document.createElement("button");
          insertButton.textContent = "Insert";
          insertButton.onclick = () => insertSingleCard(capturedRanges.length - 1 - index);

          menuBar.appendChild(infoSpan);
          menuBar.appendChild(insertButton);

          card.appendChild(rangeAddress);
          card.appendChild(menuBar);

          cardContainer.appendChild(card);
        });

      // Focus and select text of the first (most recent) input
      if (firstInput) {
        setTimeout(() => {
          firstInput.focus();
          firstInput.select();
        }, 0);
      }
    }


    function insertSingleCard(index) {
      Excel.run(async (context) => {
        const activeCell = context.workbook.getActiveCell();

        const activeSheet = context.workbook.worksheets.getActiveWorksheet();
        activeSheet.load("name"); // Explicitly load the 'name' property
        await context.sync(); // Sync to load the properties

        // Execute only if the active sheet is "Change log"
        if (activeSheet.name.toLowerCase() !== "change log") {
          console.log("Action not allowed. The active sheet is not 'Change log'.");
          showPopup("Unable to insert address", "Destination sheet must be 'Change log'.");
          return;
        }

        const data = capturedRanges[index];
        const initials = document.getElementById("initialsInput").value || "N/A";

        // Create a horizontal array with initials and address
        const valueArray = [
          [data.sheet, getAddressLinkDynamic(data.sheet, data.address, data.fullAddress), data.description, initials]
        ];

        console.log(valueArray);

        // Get a range that's two cells wide starting from the active cell
        const targetRange = activeCell.getResizedRange(0, 3);

        targetRange.load("valueTypes");

        await context.sync();

        let isEmpty = targetRange.valueTypes.every((row) =>
          row.every((cell) => cell === Excel.RangeValueType.empty || cell === null || cell === "")
        );

        if (isEmpty) {
          // Set the values of the range
          targetRange.values = valueArray;
        } else {
          showPopup("Unable to insert address", "Destination must be empty.");
        }
        targetRange.select();

        await context.sync();
      });
    }


    function getAddressLinkDynamic(sht: string, addr: string, fullAddr: string)
    {
      // Construct the LET formula with HYPERLINK
      const linkFormula = `= LET(rng, ${fullAddr}, sht, TEXTAFTER(CELL("filename", rng), "]"), addr, IF(ROWS(rng) + COLUMNS(rng)=2, ADDRESS(ROW(rng), COLUMN(rng)), ADDRESS(MIN(ROW(rng)), MIN(COLUMN(rng))) & ":" & ADDRESS(MAX(ROW(rng)), MAX(COLUMN(rng)))), dynamic_link, HYPERLINK("#'" & sht & "'!" & addr, "↗️" & SUBSTITUTE(addr, "$", "")), IFERROR(dynamic_link, HYPERLINK("#'${sht}'!${addr}","[static!] ↗️${addr.replace(
        "$",
        ""
      )}")))`;

      return linkFormula;
    }


    let popupCount = 0;


    function closePopup(popupId) {
      const popup = document.getElementById(popupId);
      if (popup) {
        popup.classList.remove("show");
        setTimeout(() => {
          popup.parentElement.remove();
        }, 300);
      }
    }


    function showPopup(title, message) {
      const container = document.getElementById("popup-container");
      const popupId = `popup-${popupCount++}`;

      // Create wrapper for positioning
      const wrapper = document.createElement("div");
      wrapper.className = "popup-wrapper";

      // Create new popup element
      const popup = document.createElement("div");
      popup.className = "popup-container";
      popup.id = popupId;

      // Add content
      popup.innerHTML = `
            <button class="popup-close" onclick="closePopup('${popupId}')">&times;</button>
            <h3 class="popup-title">${title}</h3>
            <p class="popup-message">${message}</p>
          `;

      // Add to wrapper then to container
      wrapper.appendChild(popup);
      container.appendChild(wrapper);

      // Trigger animation after a small delay to ensure proper rendering
      requestAnimationFrame(() => {
        popup.classList.add("show");
      });

      // Remove popup after animation
      setTimeout(() => {
        closePopup(popupId);
      }, 3000);
    }

    // Function to delete a card from the capturedRanges array and update the card container
    function deleteCard(index) {
      // Remove the card at the specified index from the capturedRanges array
      capturedRanges.splice(index, 1);
      // Update the card container to reflect the changes in the capturedRanges array
      updateCardContainer();
    }


    function updateDescription(index, newDescription) {
      capturedRanges[index].description = newDescription;
    }
  language: typescript
template:
  content: "<div id=\"taskpane\">\n\n\n\n\t<div class=\"button-container\">\n\t\t<button id=\"captureAddressBtn\">Capture Address</button>\n\t</div>\n\t<div class=\"input-container\">\n\t\t<label for=\"initialsInput\">Initials:</label>\n\t\t<input type=\"text\" id=\"initialsInput\" placeholder=\"XY\" />\n\t</div>\n\n\n\t\t<hr class=\"spacer\">\n\t\t<div class=\"card-container\">\n\t\t\t<div class=\"card-header\">Captured Ranges</div>\n\t\t\t<div id=\"cardContainer\"></div>\n\t\t</div>\n\n\t</div>\n\n\t<div id=\"popup-container\"></div>"
  language: html
style:
  content: |-
    body {
        font-family: Arial, sans-serif;
        margin: 0;
        padding: 10px;
    }
    #taskpane {
        width: 100%;
    }

    .main-title {
        text-align: center;
        color: #333;
    }

    .description-block {
        background-color: #f0f0f0;
        padding: 10px;
        margin-bottom: 20px;
        text-align: left;
    }

    .button-container {
        display: flex;
        justify-content: center;
        margin-bottom: 10px;
    }

    #captureAddressBtn {
        width: auto;
        padding: 10px 20px;
        font-size: 16px;
        font-weight: bold;
        color: white;
        background-color: #df4114;
        border: none;
        border-radius: 25px;
        cursor: pointer;
    }

    .input-container {
        display: flex;
        justify-content: center;
        align-items: center;
        margin-bottom: 20px;
    }

    .input-container label {
        margin-right: 10px;
    }

    #initialsInput {
        width: 50px;
        padding: 5px;
    }

    .card-container {
        background-color: #fafafa;
        padding: 10px;
        border-radius: 5px;
    }

    .card-header {
        background-color: #f0f0f0;
        padding: 5px 10px;
        font-size: 12px;
        margin-bottom: 10px;
    }

    .card {
        background-color: #f5f5f5;
        border: 1px solid #ccc;
        border-radius: 10px;
        margin-bottom: 20px;
        padding: 10px 10px 0 10px;
        overflow: hidden;
        position: relative;
        height: auto; /* Allow the card to grow with its content */
    }

    .card .range-address {
        font-size: 18px;
        font-weight: bold;
        text-align: center;
        margin: 10px 0 15px 0;
        border: none;
        background: transparent;
        width: 100%;
        transition: box-shadow 0.3s ease;
        resize: none; /* Prevents manual resizing */
        overflow: hidden; /* Hides scrollbar */
        min-height: 24px; /* Minimum height of one line */
        box-sizing: border-box;
        display: block;
        height: auto; /* Allow the textarea to grow */
    }

    .card .range-address:focus {
        outline: none;
    }
    .card .menu-bar {
        display: flex;
        justify-content: space-between;
        align-items: center;
        background-color: #e0e0e0;
        padding: 10px;
        margin: 0 -10px;
    }

    .card .menu-bar span {
        font-size: 12px;
    }

    .card .menu-bar button {
        font-size: 14px;
        font-weight: normal;
        color: #df4114;
        background-color: white;
        border: 1px solid #df4114;
        border-radius: 5px;
        padding: 6px 12px;
        cursor: pointer;
    }

    .card .menu-bar button:hover {
        background-color: #f8f8f8;
    }

    .delete-btn {
        position: absolute;
        top: 5px;
        right: 5px;
        background: none;
        border: none;
        font-size: 18px;
        cursor: pointer;
        color: #999;
    }

    .delete-btn:hover {
        color: #df4114;
    }


        #popup-container {
          position: fixed;
          bottom: 20px;
          right: 20px;
          display: flex;
          flex-direction: column-reverse;
          gap: 8px;
          z-index: 1000;
        }

        .popup-container {
          width: 300px;
          background-color: #ff0000;
          border-radius: 8px;
          padding: 16px;
          border: 2px solid #ff0000;
          opacity: 0;
          transition: opacity 0.3s ease-in-out;
        }

        .popup-container.show {
          opacity: 1;
        }

        .popup-title {
          margin: 0 0 8px 0;
          color: white;
          font-size: 16px;
          font-weight: 600;
          padding-right: 24px;
        }

        .popup-message {
          margin: 0;
          color: white;
          font-size: 14px;
          line-height: 1.5;
        }

        .popup-close {
          position: absolute;
          top: 12px;
          right: 12px;
          width: 20px;
          height: 20px;
          border: none;
          background: none;
          color: white;
          font-size: 20px;
          line-height: 1;
          cursor: pointer;
          padding: 0;
          display: flex;
          align-items: center;
          justify-content: center;
        }

        .popup-close:hover {
          opacity: 0.8;
        }


    .popup-wrapper {
      position: relative;
    }
  language: css
libraries: |
  https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js
  @types/office-js

  office-ui-fabric-core@11.1.0/dist/css/fabric.min.css
  office-ui-fabric-js@1.5.0/dist/css/fabric.components.min.css

  core-js@2.4.1/client/core.min.js
  @types/core-js

  jquery@3.1.1
  @types/jquery@3.3.1
