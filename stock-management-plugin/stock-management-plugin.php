<?php
/*
Plugin Name: Stock Management Plugin
Plugin URI: https://www.warpdevelopment.com/
Description: Custom made plugin that updates stock and generates a report based on spreadsheet data
Author: Warp Development
Version: 1.0.0
Author URI: https://www.warpdevelopment.com/
License: GPL-2.0+
License URI: http://www.gnu.org/licenses/gpl-2.0.txt
Text Domain: wpc
Requires PHP: 8.0
Requires at least: 6.3
 */
require __DIR__ . '/vendor/autoload.php';

function stock_plugin_enqueue_scripts()
{

    if (is_admin()) {

        wp_enqueue_script('jquery');
        wp_enqueue_script('plugin-script', plugin_dir_url(__FILE__) . 'plugin-script.js', array('jquery'), '1.0', true);
        wp_localize_script('plugin-script', 'stock_ajax', array(
            'ajax_url' => admin_url('admin-ajax.php'),
            'nonce' => wp_create_nonce('process_stock_speadsheet_nonce')
        ));
    }
}


add_action('admin_enqueue_scripts', 'stock_plugin_enqueue_scripts');
add_action('admin_menu', 'register_stock_upload_menu');


function register_stock_upload_menu()
{
    add_menu_page('Stock Management', 'Stock Management', 'manage_options', 'binuns-stock-management', 'my_plugin_options', 'dashicons-list-view', 55);
}


function my_plugin_options()
{
    if (!current_user_can('manage_options')) {
        wp_die(__('You do not have sufficient permissions to access this page.'));
    }
    if ($message = get_transient('form_submission_status')) {
        echo '<div class="notice notice-success is-dismissible"><p>' . esc_html($message) . '</p></div>';
        delete_transient('form_submission_status');
    }


    if (get_transient('updated_spreadsheet_path')) {
        echo '<script type="text/javascript">
                jQuery(document).ready(function($) {
                    $("#download-updated-spreadsheet").attr("href", "' . esc_url(admin_url('admin-post.php?action=download_updated_spreadsheet')) . '");
                    $("#download-button-container").show();
                });
              </script>';
    } else {
        echo '<script type="text/javascript">
                jQuery(document).ready(function($) {
                    $("#download-button-container").hide();
                });
              </script>';
    }

    echo '<style>

.tooltip {
  position: relative;
  display: inline-block;
      margin-left: 20px;
 
}
#form-messages .notice.is-dismissible, .stock-form-container .notice.notice-info,#form-messages .notice.notice-error{
    
    margin-left: 0 !important;
    margin-bottom: 15px !important;
        width: auto !important;
}

.stock-form-container .notice-success,#binuns-stock-form div.updated,.notice-info {
    border-left-color: #37665aa6 !important;
}
.stock-form-container .notice-success {
margin-bottom: 15px !important;
}
.tooltip .tooltiptext {
      visibility: hidden;
    width: 200px;
    background-color: #000000;
    color: #fff;
    text-align: center;
    padding: 13px;
    border-radius: 6px;
    position: absolute;
    z-index: 1;
    margin-left: 0px;
    top: 20px;
    font-size: 14px;
}
#download-button-container a {
    background-color: white;
    color: black;
    border-radius: 3px;    
    font-weight: 600;
    font-size: 14px;
    padding: 2px 25px;
    text-transform: uppercase;
    border: 2px solid black;
    font-family: "Open Sans", Sans-serif;
}
.tooltip:hover .tooltiptext {
  visibility: visible;
}

.stock-upload-table .col-letter-input{
    margin-left: 15px;
    width: 140px;
    height: 40px;
    font-size: 15px;
}
.stock-form-container {
    background-color: white; 
    padding: 10px 40px;
    border: 1px solid #E0E0E0;
}
h4.stock-form-title {
    font-size:24px; 
    font-family:Open-Sans, Sans-serif;
    display: flex;
    align-items: center;
    gap: 10px;
    color: black;
}
.stock-upload-table label,.stock-upload-table th {
      cursor: default;
      color: black;
}
.stock-upload-table{
    font-family: Open-Sans, sans-serif;
    text-align: left;
    font-size: 15px;
    border-spacing: 0px 10px;
    color: black;
}
.stock-upload-table td {
    display: flex;
    align-items: center;
}
.tooltip.percentage {
margin-left: 6px;
}
.stock-submit-btn {
    background-color: black;
    border-radius: 3px;
    color: white;
    font-weight: 600;
    font-size: 14px;
    padding: 10px 25px; 
    border: none;   
    font-family: "Open Sans", Sans-serif;
    text-transform: uppercase;
    cursor:pointer;
}
.upload-excel{
    width: 230px;
}
.excel-file-error,.sku-error,.qty-error,.percentage-error,.threshold-error {
font-size: 13px;
    color: #ea0003;
    display: none;
    font-family: "Open Sans", Sans-serif;
    font-weight: 700;
    margin-top: -10px !important;
}
    #clear-report-btn {
        margin-left: 20px;
    height: 37px;
    background-color: white;
    color: black;
    border: 2px solid black;
    text-transform: uppercase;
    font-weight: 700;}
</style>';
    echo '<div class="stock-form-container">';
    echo '
      <h4 class="stock-form-title">
        <img src="' . plugin_dir_url(__FILE__) . 'images/binuns-logo-reduced.png" />Binuns
        Stock Management
      </h4>';
    ?>
    <div id="form-messages"></div>
    <form id="binuns-stock-form" method="post" action="<?php echo admin_url('admin-post.php'); ?>"
          enctype="multipart/form-data">
        <input type="hidden" name="action" value="process_stock_speadsheet">
        <?php wp_nonce_field('process_stock_speadsheet_nonce'); ?>
        <div id="progress-container" style="display: none;">
    <div id="progress-bar" style="width: 0%; background-color: rgb(55, 102, 90); height: 30px; text-align: center; line-height: 30px; color: white;margin-bottom: 10px;
    margin-top: 10px;">0%</div>
</div>
        <table class="stock-upload-table">
            <tbody>
            <tr class="stock-upload">
                <th>Upload Stock Excel Spreadsheet</th>
                <td style="margin-left: 35px;" class="upload-excel">
                    <input id="excel-stock-file" type="file" name="excel-stock-file" accept=".xls, .xlsx"/>


                    <div class="tooltip">
                        <img src="<?php echo plugin_dir_url(__FILE__); ?>images/tooltip-circle.png"/>
                        <span class="tooltiptext"
                        >Please note that only Excel Spreadsheet file format (e.g. .xlsx, or .xls ) is accepted.</span
                        >
                    </div>

                </td>

            </tr>
            <tr>
                <td><p class="excel-file-error">Invalid File Type - please upload a file in the accepted .xlsx, or .xls
                        Excel Spreadsheet format.</p></td>
            </tr>
            <tr>
                <th>
                    <label for="sku-column">Column Letter for Product SKUs</label>
                </th>
                <td style="margin-left: 20px;">
                    <select id="sku-col-input" class="col-letter-input" name="sku-column">
                        <option value="" selected='selected' disabled='disabled'>Select a letter</option>

                        <option value="A">A</option>
                        <option value="B">B</option>
                        <option value="C">C</option>
                        <option value="D">D</option>
                        <option value="E">E</option>
                        <option value="F">F</option>
                        <option value="G">G</option>
                        <option value="H">H</option>
                        <option value="I">I</option>
                        <option value="J">J</option>
                        <option value="K">K</option>
                        <option value="L">L</option>
                        <option value="M">M</option>
                        <option value="N">N</option>
                        <option value="O">O</option>
                        <option value="P">P</option>
                        <option value="Q">Q</option>
                        <option value="R">R</option>
                        <option value="S">S</option>
                        <option value="T">T</option>
                        <option value="U">U</option>
                        <option value="V">V</option>
                        <option value="W">W</option>
                        <option value="X">X</option>
                        <option value="Y">Y</option>
                        <option value="Z">Z</option>
                        <option value="AA">AA</option>
                        <option value="AB">AB</option>
                        <option value="AC">AC</option>
                        <option value="AD">AD</option>
                        <option value="AE">AE</option>
                        <option value="AF">AF</option>
                        <option value="AG">AG</option>
                        <option value="AH">AH</option>
                        <option value="AI">AI</option>
                        <option value="AJ">AJ</option>
                        <option value="AK">AK</option>
                        <option value="AL">AL</option>
                        <option value="AM">AM</option>
                        <option value="AN">AN</option>
                        <option value="AO">AO</option>
                        <option value="AP">AP</option>
                        <option value="AQ">AQ</option>
                        <option value="AR">AR</option>
                        <option value="AS">AS</option>
                        <option value="AT">AT</option>
                        <option value="AU">AU</option>
                        <option value="AV">AV</option>
                        <option value="AW">AW</option>
                        <option value="AX">AX</option>
                        <option value="AY">AY</option>
                        <option value="AZ">AZ</option>

                    </select>

                    <div class="tooltip">
                        <img src="<?php echo plugin_dir_url(__FILE__); ?>images/tooltip-circle.png"/>
                        <span class="tooltiptext"
                        >Please insert the letter of the column where the SKUs of
                    the products are listed, for example "F" if the SKUs are listed in
                    Column F of the spreadsheet.</span
                        >
                    </div>
                </td>
            </tr>
            <tr>
                <td><p class="sku-error">Please enter the letter(s) of the Excel Spreadsheet column. Numbers and symbols
                        are invalid.</p></td>
            </tr>
            <tr>
                <th>
                    <label for="stock-qty-column"
                    >Column Letter for Stock Quantity</label
                    >
                </th>
                <td style="margin-left: 20px;">


                    <select id="qty-col-input" class="col-letter-input" name="stock-qty-column">
                        <option value="" selected='selected' disabled='disabled'>Select a letter</option>

                        <option value="A">A</option>
                        <option value="B">B</option>
                        <option value="C">C</option>
                        <option value="D">D</option>
                        <option value="E">E</option>
                        <option value="F">F</option>
                        <option value="G">G</option>
                        <option value="H">H</option>
                        <option value="I">I</option>
                        <option value="J">J</option>
                        <option value="K">K</option>
                        <option value="L">L</option>
                        <option value="M">M</option>
                        <option value="N">N</option>
                        <option value="O">O</option>
                        <option value="P">P</option>
                        <option value="Q">Q</option>
                        <option value="R">R</option>
                        <option value="S">S</option>
                        <option value="T">T</option>
                        <option value="U">U</option>
                        <option value="V">V</option>
                        <option value="W">W</option>
                        <option value="X">X</option>
                        <option value="Y">Y</option>
                        <option value="Z">Z</option>

                        <option value="AA">AA</option>
                        <option value="AB">AB</option>
                        <option value="AC">AC</option>
                        <option value="AD">AD</option>
                        <option value="AE">AE</option>
                        <option value="AF">AF</option>
                        <option value="AG">AG</option>
                        <option value="AH">AH</option>
                        <option value="AI">AI</option>
                        <option value="AJ">AJ</option>
                        <option value="AK">AK</option>
                        <option value="AL">AL</option>
                        <option value="AM">AM</option>
                        <option value="AN">AN</option>
                        <option value="AO">AO</option>
                        <option value="AP">AP</option>
                        <option value="AQ">AQ</option>
                        <option value="AR">AR</option>
                        <option value="AS">AS</option>
                        <option value="AT">AT</option>
                        <option value="AU">AU</option>
                        <option value="AV">AV</option>
                        <option value="AW">AW</option>
                        <option value="AX">AX</option>
                        <option value="AY">AY</option>
                        <option value="AZ">AZ</option>

                    </select>


                    <div class="tooltip">
                        <img src="<?php echo plugin_dir_url(__FILE__); ?>images/tooltip-circle.png"/>
                        <span class="tooltiptext"> Please insert the letter of the column where the stock
                            quantities of the products are listed, for example "A" if the stock quantities are found in Column A of the spreadsheet.</span>
                    </div>
                </td>
            </tr>
            <tr>
                <td><p class="qty-error">Please enter the letter(s) of the Excel Spreadsheet Column. Numbers and symbols
                        are invalid.</p></td>
            </tr>
            <tr>
                <th>
                    <label for="stock-percentage-column"
                    >Enter the percentage of stock to be updated</label>
                </th>
                <td style="margin-left: 20px;">
                    <input
                            class="col-letter-input"
                            id="percentage-input"
                            type="number"
                            step="1"
                            min="0"
                            max="100"
                            name="stock-percentage-column"
                    /><span>%</span>
                    <div class="tooltip percentage">
                        <img src="<?php echo plugin_dir_url(__FILE__); ?>images/tooltip-circle.png"
                        /><span class="tooltiptext"
                        >For example, if a product as per the spreadsheet has 100 units in stock and 80% has been inputted, then only 80 units of stock will be loaded onto the website. </span>
                    </div>
                </td>
            </tr>
            <tr>
                <td><p class="percentage-error">Please enter a value between 0 - 100.</p></td>
            </tr>
            <tr>
                <th>
                    <label for="stock-threshold-column"
                    >Enter the minimum stock threshold</label>
                </th>
                <td style="margin-left: 20px;">
                    <input
                            id="threshold-input"
                            class="col-letter-input"
                            type="number"
                            name="stock-threshold-column"
                            step="1"
                            min="0"
                    />
                    <div class="tooltip">
                        <img src="<?php echo plugin_dir_url(__FILE__); ?>images/tooltip-circle.png"
                        /><span class="tooltiptext"
                        >If the calculated stock quantity from the above is below this specified minimum threshold, the stock quantity will be adjusted to meet this minimum threshold instead.</span>
                    </div>
                </td>
            </tr>
            <tr>
                <td><p class="threshold-error">Please enter a value.</p></td>
            </tr>
            <tr style="    display: flex;
    gap: 20px;">

                <td style="margin-top: 25px;">
                    <input
                            class="stock-submit-btn"
                            type="submit"
                            name="stock-submit-btn"
                    />

                </td>
                <td style="margin-top: 25px;">
                    <div id="download-button-container" style="display: none;">
                        <a href="#" class="button button-primary" id="download-updated-spreadsheet">Download
                            Spreadsheet</a><button id="clear-report-btn" class="button" type="button">Clear</button>   <img id="loader-stock" src="/wp-content/uploads/2023/11/loader.gif"  style="display: none;width: 20px;position: relative; left: 15px;top: 5px;"></div>

     
      
                </td>
            </tr>
            </tbody>
        </table>


    </form>
    <?php
echo '<script type="text/javascript">
jQuery(document).ready(function($) {
    $("#clear-report-btn").on("click", function() {
    $("#loader-stock").show();
        $.ajax({
            url: ajaxurl,
            type: "POST",
            data: {
                action: "clear_stock_report"
            },
            success: function(response) {
                if (response.success) {
                    location.reload();
                } else {
                    alert("Failed to clear report.");
                $("#loader-stock").hide();
                }
            }
        });
    });
});
</script>';
    echo '</div>';


}


function process_stock_speadsheet() {
    if (!check_ajax_referer('process_stock_speadsheet_nonce', 'security', false)) {
        wp_send_json_error('Security check failed');
        return;
    }

    if (!isset($_FILES['excel-stock-file'])) {
        wp_send_json_error('No file uploaded');
        return;
    }

    $file = $_FILES['excel-stock-file'];
    $upload_dir = wp_upload_dir();
    $file_name = sanitize_file_name($file['name']);
    $file_path = $upload_dir['path'] . '/' . $file_name;

    if (move_uploaded_file($file['tmp_name'], $file_path)) {
        require_once ABSPATH . 'wp-admin/includes/file.php';
        WP_Filesystem();
        global $wp_filesystem;

        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($file_path);
        $worksheet = $spreadsheet->getActiveSheet();
        $highestRow = $worksheet->getHighestRow();

        $skuColumn = sanitize_text_field($_POST['sku-column']);
        $qtyColumn = sanitize_text_field($_POST['stock-qty-column']);
        $percentage = floatval($_POST['stock-percentage-column']) / 100;
        $thresholdValue = isset($_POST['stock-threshold-column']) ? intval($_POST['stock-threshold-column']) : 0;

        $chunkSize = 100; // Adjust based on your needs
        $totalChunks = ceil(($highestRow - 1) / $chunkSize);
        $highestColumn = $worksheet->getHighestColumn();
        $highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn);

        // Safely set new columns after the highest column in the spreadsheet
        $maxAllowedColumnIndex = 16384; // Excel's max column limit (XFD)
        if ($highestColumnIndex + 4 > $maxAllowedColumnIndex) {
            throw new Exception('Invalid cell coordinate: one of the required columns exceeds the spreadsheet size.');
        }

        $preUpdateQtyCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($highestColumnIndex + 1);
        $postUpdateQtyCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($highestColumnIndex + 2);
        $thresholdCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($highestColumnIndex + 3);
        $productExistsCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($highestColumnIndex + 4);

        // Now safely continue to set values in these columns
        $worksheet->setCellValue($preUpdateQtyCol . '1', 'Quantity Pre Update');
        $worksheet->setCellValue($postUpdateQtyCol . '1', 'Quantity Post Update');
        $worksheet->setCellValue($thresholdCol . '1', 'Product on minimum threshold?');
        $worksheet->setCellValue($productExistsCol . '1', 'Product exists on the system');

        $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
        $temp_file = tempnam(sys_get_temp_dir(), 'updated_stock_') . '.xlsx';
        $writer->save($temp_file);

        set_transient('updated_spreadsheet_path', $temp_file, DAY_IN_SECONDS);
        set_transient('stock_update_progress', 0, DAY_IN_SECONDS);
        set_transient('stock_update_total_chunks', $totalChunks, DAY_IN_SECONDS);

        wp_send_json_success([
            'message' => 'File uploaded and prepared successfully',
            'filePath' => $temp_file,
            'totalChunks' => $totalChunks,
            'chunkSize' => $chunkSize,
            'totalRows' => $highestRow - 1,
            'skuColumn' => $skuColumn,
            'qtyColumn' => $qtyColumn,
            'percentage' => $percentage,
            'thresholdValue' => $thresholdValue,
            'preUpdateQtyCol' => $preUpdateQtyCol,
            'postUpdateQtyCol' => $postUpdateQtyCol,
            'thresholdCol' => $thresholdCol,
            'productExistsCol' => $productExistsCol
        ]);
    } else {
        wp_send_json_error('Failed to move uploaded file');
    }
}
add_action('wp_ajax_process_stock_speadsheet', 'process_stock_speadsheet');

function process_stock_batch() {
    if (!check_ajax_referer('process_stock_speadsheet_nonce', 'security', false)) {
        wp_send_json_error('Security check failed');
        return;
    }

    $chunkNumber = isset($_POST['chunkNumber']) ? intval($_POST['chunkNumber']) : 0;
    $filePath = isset($_POST['filePath']) ? $_POST['filePath'] : '';
    $chunkSize = isset($_POST['chunkSize']) ? intval($_POST['chunkSize']) : 100;
    $skuColumn = isset($_POST['skuColumn']) ? $_POST['skuColumn'] : 'A';
    $qtyColumn = isset($_POST['qtyColumn']) ? $_POST['qtyColumn'] : 'B';
    $percentage = isset($_POST['percentage']) ? floatval($_POST['percentage']) : 1;
    $thresholdValue = isset($_POST['thresholdValue']) ? intval($_POST['thresholdValue']) : 0;
    $totalChunks = isset($_POST['totalChunks']) ? intval($_POST['totalChunks']) : 1;
    $preUpdateQtyCol = isset($_POST['preUpdateQtyCol']) ? $_POST['preUpdateQtyCol'] : '';
    $postUpdateQtyCol = isset($_POST['postUpdateQtyCol']) ? $_POST['postUpdateQtyCol'] : '';
    $thresholdCol = isset($_POST['thresholdCol']) ? $_POST['thresholdCol'] : '';
    $productExistsCol = isset($_POST['productExistsCol']) ? $_POST['productExistsCol'] : '';

    if (!$filePath || !file_exists($filePath)) {
        wp_send_json_error('Invalid file path');
        return;
    }

    try {
        $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
        $reader->setReadDataOnly(false);
        $spreadsheet = $reader->load($filePath);
        $worksheet = $spreadsheet->getActiveSheet();
        
        $startRow = ($chunkNumber - 1) * $chunkSize + 2;
        $endRow = min($startRow + $chunkSize - 1, $worksheet->getHighestRow());

        $highestColumn = $worksheet->getHighestColumn();
        $highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn);

        // Ensure new columns do not exceed the max allowed columns
        $maxAllowedColumnIndex = 16384; // Excel's max column limit (XFD)
        if ($highestColumnIndex + 4 > $maxAllowedColumnIndex) {
            throw new Exception('Invalid cell coordinate: one of the required columns exceeds the spreadsheet size.');
        }

        $skuQtyMap = [];
        $updatedProducts = 0;
        $processedProducts = 0;

        for ($row = $startRow; $row <= $endRow; $row++) {
            $sku = $worksheet->getCell($skuColumn . $row)->getValue();
            $qty = $worksheet->getCell($qtyColumn . $row)->getValue();

            if ($sku && is_numeric($qty)) {
                $adjustedQty = floor(floatval($qty) * $percentage);
                $skuQtyMap[$sku] = [
                    'quantity' => $adjustedQty,
                    'row' => $row,  // Store the row number for later use
                ];
                $processedProducts++;
            } else {
                // Mark invalid products in the spreadsheet
                $worksheet->setCellValue($preUpdateQtyCol . $row, 'N/A');
                $worksheet->setCellValue($postUpdateQtyCol . $row, 'N/A');
                $worksheet->setCellValue($thresholdCol . $row, 'N/A');
                $worksheet->setCellValue($productExistsCol . $row, 'No');
            }
        }

        // Perform bulk update for this chunk
        $result = bulk_update_product_stock($skuQtyMap, $thresholdValue);

        // Update the spreadsheet with the bulk update results
        foreach ($result as $sku => $data) {
            $row = $skuQtyMap[$sku]['row'];
        
            // Ensure values are scalar types before setting cell values
            $preUpdateQty = is_scalar($data['preUpdateQty']) ? $data['preUpdateQty'] : 'N/A';
            $postUpdateQty = is_scalar($data['postUpdateQty']) ? $data['postUpdateQty'] : 'N/A';
            $thresholdUsed = is_scalar($data['thresholdUsed']) ? ($data['thresholdUsed'] ? 'Yes' : 'No') : 'N/A';
            $productExists = is_scalar($data['productExists']) ? ($data['productExists'] ? 'Yes' : 'No') : 'N/A';
        
            // Set the values in the spreadsheet cells
            $worksheet->setCellValue($preUpdateQtyCol . $row, $preUpdateQty);
            $worksheet->setCellValue($postUpdateQtyCol . $row, $postUpdateQty);
            $worksheet->setCellValue($thresholdCol . $row, $thresholdUsed);
            $worksheet->setCellValue($productExistsCol . $row, $productExists);
        
            if ($data['productExists'] && $data['updated']) {
                $updatedProducts++;
            }
        }

        // Save the updated spreadsheet
        $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save($filePath);

        // Calculate progress
        $progress = ($chunkNumber / $totalChunks) * 100;
        set_transient('stock_update_progress', $progress, DAY_IN_SECONDS);

        // Return the response
        wp_send_json_success([
            'chunkNumber' => $chunkNumber,
            'updatedProducts' => $updatedProducts,
            'processedProducts' => $processedProducts,
            'progress' => $progress
        ]);

    } catch (Exception $e) {
        // Log the error and return the error response
        error_log('Error processing chunk: ' . $e->getMessage());
        wp_send_json_error('Error processing chunk: ' . $e->getMessage());
    }
}
add_action('wp_ajax_process_stock_batch', 'process_stock_batch');
function bulk_update_product_stock($skuQtyMap, $threshold) {
    global $wpdb;
    $skuCases = [];
    $ids = [];
    $result = [];

    foreach ($skuQtyMap as $sku => $data) {
        $product_id = wc_get_product_id_by_sku($sku);
        $quantity = $data['quantity'];
        $row = $data['row']; // Row number from spreadsheet

        if ($product_id) {
            // Get the current stock quantity
            $preUpdateQty = get_post_meta($product_id, '_stock', true);
            $newQuantity = ($quantity <= $threshold) ? 0 : $quantity;

            // Prepare bulk update query
            $skuCases[] = $wpdb->prepare("WHEN %d THEN %d", $product_id, $newQuantity);
            $ids[] = $product_id;

            // Prepare result array for spreadsheet update
            $result[$sku] = [
                'preUpdateQty' => $preUpdateQty,
                'postUpdateQty' => $newQuantity,
                'thresholdUsed' => ($quantity <= $threshold),
                'productExists' => true,
                'updated' => ($preUpdateQty != $newQuantity)
            ];
        } else {
            // Product doesn't exist, mark as such in the result
            $result[$sku] = [
                'preUpdateQty' => 'N/A',
                'postUpdateQty' => 'N/A',
                'thresholdUsed' => false,
                'productExists' => false,
                'updated' => false
            ];
        }
    }

    // Only run the bulk update query if there are cases to update
    if (!empty($skuCases)) {
        $sql = "
            UPDATE {$wpdb->prefix}postmeta
            SET meta_value = CASE post_id
                " . implode(' ', $skuCases) . "
                END
            WHERE meta_key = '_stock' AND post_id IN (" . implode(',', $ids) . ")
        ";
        $wpdb->query($sql);
    }

    return $result;
}
function update_product_stock($sku, $quantity, $threshold) {
    $product_id = wc_get_product_id_by_sku($sku);
    if (!$product_id) {
        error_log("Product with SKU $sku not found");
        return array('productExists' => false);
    }
    $product = wc_get_product($product_id);
    $preUpdateQty = $product->get_stock_quantity();
    $thresholdUsed = false;

    // Initialize newQuantity
    $newQuantity = $preUpdateQty;

    // Apply threshold logic: Set stock to 0 if quantity is 0 or less than/equal to the threshold
    if ($quantity == 0 || $quantity <= $threshold) {
        $thresholdUsed = true;
        $newQuantity = 0;
    } else {
        // If quantity is above the threshold, set it to the floored value of the quantity
        $newQuantity = floor($quantity);
    }

    // Only update if the stock quantity is different
    if ($preUpdateQty !== $newQuantity) {
        $product->set_stock_quantity($newQuantity);
        $product->save();
        error_log("Updated product $sku: $preUpdateQty -> $newQuantity (Threshold: $threshold, Original Quantity: $quantity)");
    } else {
        error_log("No change in quantity for product $sku (Threshold: $threshold, Original Quantity: $quantity)");
    }

    // Return the update result
    return array(
        'preUpdateQty' => $preUpdateQty,
        'postUpdateQty' => $newQuantity,
        'thresholdUsed' => $thresholdUsed,
        'productExists' => true,
        'updated' => ($preUpdateQty !== $newQuantity)
    );
}
function display_form_submission_message()
{
    if ($message = get_transient('form_submission_status')) {

        $class = strpos($message, 'Invalid') === 0 ? 'notice-error' : 'notice-success';

        echo '<script type="text/javascript">
                document.addEventListener("DOMContentLoaded", function() {
                    var messageHtml = "<div class=\'notice ' . $class . ' is-dismissible\'><p>' . esc_html($message) . '</p></div>";
                    document.getElementById("form-messages").innerHTML = messageHtml;
                });
              </script>';

        delete_transient('form_submission_status');
    }
}

add_action('admin_notices', 'display_form_submission_message');


function handle_download_updated_spreadsheet() {
    check_ajax_referer('process_stock_speadsheet_nonce', 'security');
    
    $filePath = get_transient('updated_spreadsheet_path');
    if ($filePath && file_exists($filePath)) {
        $fileName = basename($filePath);
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="' . $fileName . '"');
        header('Content-Length: ' . filesize($filePath));
        header('Cache-Control: max-age=0');
        header('Pragma: public');
        
        readfile($filePath);
        exit;
    } else {
        wp_send_json_error('File not found or access denied.');
    }
}
add_action('wp_ajax_download_updated_spreadsheet', 'handle_download_updated_spreadsheet');

function clear_stock_report() {
    $file_path = get_transient('updated_spreadsheet_path');
    if ($file_path && file_exists($file_path)) {
        unlink($file_path);
    }
    delete_transient('updated_spreadsheet_path');
    wp_send_json_success('Report cleared successfully.');
}
add_action('wp_ajax_clear_stock_report', 'clear_stock_report');

function check_file_permissions($filePath) {
    if (!file_exists($filePath)) {
        error_log("File does not exist: $filePath");
        return false;
    }
    if (!is_readable($filePath)) {
        error_log("File is not readable: $filePath");
        return false;
    }
    return true;
}

function check_file_integrity($filePath) {
    $zip = new ZipArchive;
    if ($zip->open($filePath) === TRUE) {
        $zip->close();
        return true;
    } else {
        error_log("Failed to open Excel file as ZIP: $filePath");
        return false;
    }
}

function check_stock_update_progress() {
    $progress = get_transient('stock_update_progress');
    $totalChunks = get_transient('stock_update_total_chunks');
    $currentChunk = ceil(($progress / 100) * $totalChunks);
    
    wp_send_json_success([
        'progress' => $progress,
        'currentChunk' => $currentChunk,
        'totalChunks' => $totalChunks
    ]);
}

add_action('wp_ajax_check_stock_update_progress', 'check_stock_update_progress');