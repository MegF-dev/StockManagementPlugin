jQuery(document).ready(function ($) {
    $("#excel-stock-file").on("change", function () {
        var fileType = this.files.length > 0 ? this.files[0].type : "";
        var validTypes = [
            "application/vnd.ms-excel", // for .xls files
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", // for .xlsx files
        ];
        if (validTypes.includes(fileType)) {
            $(".excel-file-error").hide();
        } else {
            $(".excel-file-error")
                .text(
                    "Invalid file type. Please upload an Excel Spreadsheet file."
                )
                .show();
        }
    });

    var globalFilePath = "";

    $("#binuns-stock-form").on("submit", function (e) {
        e.preventDefault();
        console.log("Form submitted");

        var isValid = true;

        var fileInput = $("#excel-stock-file")[0];
        if (fileInput.files.length === 0) {
            $(".excel-file-error").text("Please upload a file.").show();
            isValid = false;
        } else {
            var fileType = fileInput.files[0].type;
            var validTypes = [
                "application/vnd.ms-excel", // for .xls files
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", // for .xlsx files
            ];
            if (!validTypes.includes(fileType)) {
                $(".excel-file-error")
                    .text(
                        "Invalid file type. Please upload an Excel Spreadsheet file."
                    )
                    .show();
                isValid = false;
            } else {
                $(".excel-file-error").hide();
            }
        }
        var skuValue = $("#sku-col-input").val();
        if (!skuValue) {
            $(".sku-error")
                .text("Please select a letter for the SKU column.")
                .show();
            isValid = false;
        } else {
            $(".sku-error").hide();
        }

        var qtyValue = $("#qty-col-input").val();
        if (!qtyValue) {
            $(".qty-error")
                .text("Please select a letter for the Quantity column.")
                .show();
            isValid = false;
        } else {
            $(".qty-error").hide();
        }

        var percentageValue = $("#percentage-input").val().trim();
        if (!percentageValue) {
            $(".percentage-error")
                .text("Percentage input cannot be empty.")
                .show();
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
        var minThresholdValue = $("#threshold-input").val().trim();
        if (!minThresholdValue) {
            $(".threshold-error")
                .text("Minimum threshold input cannot be empty.")
                .show();
            isValid = false;
        } else {
            $(".threshold-error").hide();
        }

        if (isValid) {
            var processingDiv = $(
                '<div class="notice notice-info"></div>'
            ).html(
                "<p>Processing started... this could take several minutes. Please do not close the browser.</p>"
            );
            $("#binuns-stock-form").prepend(processingDiv);

            var formData = new FormData(this);
            formData.append("action", "process_stock_speadsheet");
            formData.append("security", stock_ajax.nonce);

            $.ajax({
                url: ajaxurl,
                type: "POST",
                data: formData,
                processData: false,
                contentType: false,
                success: function (response) {
                    console.log("Full response:", response);
                    if (response.success) {
                        console.log("Data received:", response.data);
                        $("#form-messages").html(
                            '<div style="margin-left:0 !important" class="notice notice-success file-upload-notice"><p>File uploaded successfully. Starting processing...</p></div>'
                        );

                        startChunkedProcessing(response.data);
                    } else {
                        console.error("Error in response:", response.data);
                        $("#form-messages").html(
                            '<div class="notice notice-error"><p>Error: ' +
                                response.data +
                                "</p></div>"
                        );
                        $("#binuns-stock-form .notice.notice-info").hide();
                    }
                },
                error: function (jqXHR, textStatus, errorThrown) {
                    console.error("AJAX Error:", textStatus, errorThrown);
                    $("#form-messages").html(
                        '<div class="notice notice-error"><p>An error occurred. Please try again.</p></div>'
                    );
                    $("#binuns-stock-form .notice.notice-info").hide();
                },
            });
        }
    });
    function startChunkedProcessing(data) {
        var currentChunk = 1;
        var totalUpdated = 0;

        $("#progress-container").show();
        updateProgress(0, data.totalChunks, 0, data.totalRows);

        function processChunk() {
            $.ajax({
                url: ajaxurl,
                type: "POST",
                data: {
                    action: "process_stock_batch",
                    security: stock_ajax.nonce,
                    chunkNumber: currentChunk,
                    totalChunks: data.totalChunks,
                    filePath: data.filePath,
                    chunkSize: data.chunkSize,
                    skuColumn: data.skuColumn,
                    qtyColumn: data.qtyColumn,
                    percentage: data.percentage,
                    thresholdValue: data.thresholdValue,
                    preUpdateQtyCol: data.preUpdateQtyCol,
                    postUpdateQtyCol: data.postUpdateQtyCol,
                    thresholdCol: data.thresholdCol,
                    productExistsCol: data.productExistsCol,
                },
                success: function (response) {
                    if (response.success) {
                        totalUpdated += response.data.updatedProducts;
                        updateProgress(
                            currentChunk,
                            data.totalChunks,
                            totalUpdated,
                            data.totalRows
                        );

                        currentChunk++;
                        if (currentChunk <= data.totalChunks) {
                            setTimeout(processChunk, 1000); // 1 second delay between chunks
                        } else {
                            finishProcessing(
                                totalUpdated,
                                data.totalRows,
                                data.filePath
                            );
                        }
                    } else {
                        console.error("Chunk processing error:", response.data);
                        $("#form-messages").html(
                            '<div class="notice notice-error"><p>Error processing chunk: ' +
                                response.data +
                                "</p></div>"
                        );
                        $("#binuns-stock-form .notice.notice-info").hide();
                    }
                },
                error: function (jqXHR, textStatus, errorThrown) {
                    console.error("AJAX Error:", textStatus, errorThrown);
                    $("#form-messages").html(
                        '<div class="notice notice-error"><p>An error occurred during chunk processing. Please try again.</p></div>'
                    );
                    $("#binuns-stock-form .notice.notice-info").hide();
                },
            });
        }

        processChunk();
    }

    function updateProgress(
        currentChunk,
        totalChunks,
        updatedProducts,
        totalRows
    ) {
        var percentage = Math.round((currentChunk / totalChunks) * 100);
        $("#progress-bar")
            .css("width", percentage + "%")
            .text(percentage + "%");
        $("#progress-info").text(
            "Updated " +
                updatedProducts +
                " out of " +
                totalRows +
                " products (Chunk " +
                currentChunk +
                " of " +
                totalChunks +
                ")"
        );
    }

    function finishProcessing(totalUpdated, totalRows, finalFilePath) {
        $("#progress-container").hide();
        $("#form-messages").html(
            '<div style="margin-left:0 !important" class="notice notice-success"><p>Processing complete! Updated ' +
                totalUpdated +
                " out of " +
                totalRows +
                " products.</p></div>"
        );
        $("#binuns-stock-form .notice.notice-info").hide();
        $("#download-button-container").show();

        // Update global file path with the final file path
        globalFilePath = finalFilePath;
    }

    $("#download-updated-spreadsheet").on("click", function (e) {
        e.preventDefault();
        $.ajax({
            url: ajaxurl,
            type: "POST",
            data: {
                action: "download_updated_spreadsheet",
                security: stock_ajax.nonce,
            },
            xhrFields: {
                responseType: "blob",
            },
            success: function (blob, status, xhr) {
                if (blob instanceof Blob) {
                    var downloadUrl = URL.createObjectURL(blob);
                    var a = document.createElement("a");
                    a.style.display = "none";
                    a.href = downloadUrl;
                    a.download = "updated_stock_report.xlsx";
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(downloadUrl);
                    document.body.removeChild(a);
                } else {
                    var reader = new FileReader();
                    reader.onload = function () {
                        console.error("Server response:", reader.result);
                        alert(
                            "Failed to download the file. Server response was not as expected."
                        );
                    };
                    reader.readAsText(blob);
                }
            },
            error: function (jqXHR, textStatus, errorThrown) {
                console.error("Download failed:", textStatus, errorThrown);
                alert(
                    "Failed to download the file. Please check the console for more information."
                );
            },
        });
    });

    $("#clear-report-btn").on("click", function () {
        $("#loader-stock").show();
        $.ajax({
            url: ajaxurl,
            type: "POST",
            data: {
                action: "clear_stock_report",
                security: stock_ajax.nonce,
            },
            success: function (response) {
                if (response.success) {
                    location.reload();
                } else {
                    alert("Failed to clear report.");
                    $("#loader-stock").hide();
                }
            },
        });
    });
});
