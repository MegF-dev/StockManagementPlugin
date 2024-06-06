jQuery(document).ready(function ($) {
  $("#excel-stock-file").on("change", function () {
    var fileType = this.files.length > 0 ? this.files[0].type : "";
    var validTypes = [
      "application/vnd.ms-excel", // for .xls files
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", // for .xlsx files
    ];
    if (validTypes.includes(fileType)) {
      $(".excel-file-error").hide(); // Hide error message if file type is valid
    } else {
      $(".excel-file-error")
        .text("Invalid file type. Please upload an Excel Spreadsheet file.")
        .show();
    }
  });

  $("#binuns-stock-form").on("submit", function (e) {
    var isValid = true;

    var fileInput = $("#excel-stock-file")[0];
    if (fileInput.files.length > 0) {
      var fileType = fileInput.files[0].type;
      var validTypes = [
        "application/vnd.ms-excel", // for .xls files
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", // for .xlsx files
      ];
      if (!validTypes.includes(fileType)) {
        $(".excel-file-error")
          .text("Invalid file type. Please upload an Excel Spreadsheet file.")
          .show();
        isValid = false;
      } else {
        $(".excel-file-error").hide();
      }
    } else {
      $(".excel-file-error").text("Please select a file.").show();
      isValid = false;
    }

    var skuValue = $("#sku-col-input").val();
    if (!skuValue) {
      $(".sku-error").text("Please select a letter for the SKU column.").show();
      isValid = false;
    } else {
      $(".sku-error").hide();
    }

    // Validate Quantity column input
    var qtyValue = $("#qty-col-input").val();
    if (!qtyValue) {
      $(".qty-error")
        .text("Please select a letter for the Quantity column.")
        .show();
      isValid = false;
    } else {
      $(".qty-error").hide();
    }

    // Validate Percentage input
    var percentageValue = $("#percentage-input").val().trim();
    if (!percentageValue) {
      $(".percentage-error").text("Percentage input cannot be empty.").show();
      isValid = false;
    } else if (
      isNaN(percentageValue) ||
      percentageValue < 0 ||
      percentageValue > 100
    ) {
      $(".percentage-error")
        .text("Percentage must be a number between 0 and 100.")
        .show();
      isValid = false;
    } else {
      $(".percentage-error").hide();
    }
    var minThresholdValue = $("#threshold-input").val().trim(); // Ensure we trim whitespace
    if (!minThresholdValue) {
      $(".threshold-error")
        .text("Minimum threshold input cannot be empty.")
        .show();
      isValid = false;
    } else {
      // Optionally, you can add more validation here if needed (e.g., check if it's a number)
      $(".threshold-error").hide();
    }
    // Show processing message only if all inputs are valid
    if (isValid) {
      var processingDiv = $('<div class="notice notice-info"></div>').html(
        "<p>Processing form... this could take several minutes. Please do not close the browser.</p>"
      );
      $("#binuns-stock-form").prepend(processingDiv);
    } else {
      e.preventDefault(); // Prevent form submission if not valid
    }
  });
});
