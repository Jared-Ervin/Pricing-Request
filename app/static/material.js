window.onload = function() {
  materializeControls();
};
function materializeControls() {
  materializeTextInputs();
  materializeSelects();
  materializeRadioButtons();
  materializeCheckboxes();
  materializeTables();
  materializeLists();
  materializeButtons();
}

function materializeTextInputs() {
  var label, parentEl;
  document
    .querySelectorAll('input[type="text"], textarea')
    .forEach(function(control) {
      parentEl = control.parentElement;
      control.classList.add("mdl-textfield__input");
      if (parentEl.tagName !== "DIV") {
        return;
      }
      parentEl.classList.add("mdl-textfield", "mdl-js-textfield");
      label = parentEl.querySelector("label");
      if (label) {
        label.setAttribute("for", control.id || control.name);
        label.classList.add("mdl-textfield__label");
      }
    });
}

function materializeSelects() {
  var label, parentEl;
  document.querySelectorAll("select").forEach(function(control) {
    parentEl = control.parentElement;
    control.classList.add("mdl-selectfield__select");
    if (parentEl.tagName !== "DIV") {
      return;
    }
    parentEl.classList.add("mdl-selectfield", "mdl-js-selectfield");
    label = parentEl.querySelector("label");
    if (label) {
      label.setAttribute("for", control.id || control.name);
      label.classList.add("mdl-selectfield__label");
    }
  });
}

function materializeRadioButtons() {
  var parentEl;
  document.querySelectorAll('input[type="radio"]').forEach(function(control) {
    parentEl = control.parentElement;
    control.classList.add("mdl-radio__button");
    if (parentEl.tagName == "LABEL") {
      parentEl.setAttribute("for", control.id || control.name);
      parentEl.classList.add(
        "mdl-radio",
        "mdl-js-radio",
        "mdl-js-ripple-effect"
      );
    }
  });
}

function materializeCheckboxes() {
  var parentEl;
  document
    .querySelectorAll('input[type="checkbox"]')
    .forEach(function(control) {
      parentEl = control.parentElement;
      control.classList.add("mdl-checkbox__input");
      if (parentEl.tagName == "LABEL") {
        parentEl.setAttribute("for", control.id || control.name);
        parentEl.classList.add(
          "mdl-checkbox",
          "mdl-js-checkbox",
          "mdl-js-ripple-effect"
        );
      }
    });
}

function materializeButtons() {
  document.querySelectorAll("button").forEach(function(control) {
    control.classList.add(
      "mdl-button",
      "mdl-js-button",
      "mdl-button--raised",
      "mdl-js-ripple-effect",
      "mdl-button--colored"
    );
  });
}

function materializeTables() {
  document.querySelectorAll("table").forEach(function(table) {
    table.classList.add("mdl-data-table", "mdl-js-data-table");
    table.querySelectorAll("th,td").forEach(function(cell) {
      cell.classList.add("mdl-data-table__cell--non-numeric");
    });
  });
}

function materializeLists() {
  document.querySelectorAll("ul").forEach(function(ulEl) {
    ulEl.classList.add("mdl-list");
    ulEl.querySelectorAll("li").forEach(function(liEl) {
      liEl.classList.add("mdl-list__item");
      liEl.innerHTML =
        "<span class='mdl-list__item-primary-content'>" +
        "<i class='material-icons mdl-list__item-icon'>home</i>" +
        liEl.innerText +
        "</span>";
    });
  });
}
