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


    if ($updatedFile = get_transient('updated_spreadsheet_path')) {
        echo '<script type="text/javascript">
                jQuery(document).ready(function($) {
                    $("#download-updated-spreadsheet").attr("href", "' . esc_url(admin_url('admin-post.php?action=download_updated_spreadsheet')) . '");
                    $("#download-button-container").show();
                });
              </script>';
    }

    echo '<style>

.tooltip {
  position: relative;
  display: inline-block;
      margin-left: 20px;
 
}
#form-messages .notice.is-dismissible, .stock-form-container .notice.notice-info {
    
    margin-left: 0 !important;
    margin-bottom: 15px !important;
    width: fit-content !important;
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
                            Spreadsheet</a></div>

                </td>
            </tr>
            </tbody>
        </table>


    </form>
    <?php

    echo '</div>';


}


function process_stock_speadsheet()
{

    if (!isset($_POST['_wpnonce']) || !wp_verify_nonce($_POST['_wpnonce'], 'process_stock_speadsheet_nonce')) {
        set_transient('form_submission_status', 'Security check failed', 10);
        wp_redirect(admin_url('admin.php?page=binuns-stock-management'));
        exit;
    }

    if (isset($_FILES['excel-stock-file'])) {
        $file = $_FILES['excel-stock-file'];
        $file_type = $file['type'];


        $allowed_types = array(
            'application/vnd.ms-excel',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );

        if (!in_array($file_type, $allowed_types)) {
            set_transient('form_submission_status', 'Invalid file type. Please upload an Excel Spreadsheet file.', 10);
        } else {


            $skuColumn = sanitize_text_field($_POST['sku-column']);
            $qtyColumn = sanitize_text_field($_POST['stock-qty-column']);
            $percentage = floatval($_POST['stock-percentage-column']) / 100;

            $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($file['tmp_name']);
            $spreadsheet = $reader->load($file['tmp_name']);
            $worksheet = $spreadsheet->getActiveSheet();

            $highestColumn = $worksheet->getHighestColumn();
            $highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn);

            $preUpdateQtyCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($highestColumnIndex + 1);
            $postUpdateQtyCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($highestColumnIndex + 2);
            $thresholdCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($highestColumnIndex + 3);
            $productExistsCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($highestColumnIndex + 4);

            $worksheet->setCellValue($preUpdateQtyCol . '1', 'Quantity Pre Update');
            $worksheet->setCellValue($postUpdateQtyCol . '1', 'Quantity Post Update');
            $worksheet->setCellValue($thresholdCol . '1', 'Product on minimum threshold?');
            $worksheet->setCellValue($productExistsCol . '1', 'Product exists on the system');


            foreach ($worksheet->getRowIterator(2) as $row) {
                $rowIndex = $row->getRowIndex();
                $skuCell = $worksheet->getCell($skuColumn . $row->getRowIndex());
                $qtyCell = $worksheet->getCell($qtyColumn . $row->getRowIndex());

                $sku = $skuCell->getValue();
                $qty = floatval($qtyCell->getValue()) * $percentage;


                $updateResult = update_product_stock($sku, $qty);

               

                $worksheet->setCellValue($preUpdateQtyCol . $rowIndex, $updateResult['preUpdateQty']);
                $worksheet->setCellValue($postUpdateQtyCol . $rowIndex, floor($updateResult['postUpdateQty']));
                $worksheet->setCellValue($thresholdCol . $rowIndex, $updateResult['thresholdUsed'] ? 'Yes' : 'No');
                $worksheet->setCellValue($productExistsCol . $rowIndex, $updateResult['productExists'] ? 'Yes' : 'No');
            }
            $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
            $temp_file = tempnam(sys_get_temp_dir(), 'updated_stock_') . '.xlsx';
            $writer->save($temp_file);

            set_transient('updated_spreadsheet_path', $temp_file, 60);
            set_transient('form_submission_status', 'Product stock has been successfully updated.', 60);

        }

        wp_redirect(admin_url('admin.php?page=binuns-stock-management'));
        exit;
    }

}

add_action('admin_post_process_stock_speadsheet', 'process_stock_speadsheet');

function update_product_stock($sku, $quantity)
{
    $product_id = wc_get_product_id_by_sku($sku);
    $thresholdUsed = false;
    $preUpdateQty = 0;
    $postUpdateQty = 0;
    $productExists = false;

    if ($product_id) {
        $productExists = true;
        $product = wc_get_product($product_id);
        $preUpdateQty = $product->get_stock_quantity();

        $threshold = isset($_POST['stock-threshold-column']) ? intval($_POST['stock-threshold-column']) : 0;
        if ($quantity < $threshold) {
            $product->set_stock_quantity($threshold);
            $thresholdUsed = true;
            $postUpdateQty = $threshold;
        } else {
            $product->set_stock_quantity($quantity);
            $postUpdateQty = $quantity;
        }

        $product->save();
    }

    return array('preUpdateQty' => $preUpdateQty, 'thresholdUsed' => $thresholdUsed, 'postUpdateQty' => $postUpdateQty, 'productExists' => $productExists);
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


function handle_download_updated_spreadsheet()
{
    if ($filePath = get_transient('updated_spreadsheet_path')) {
        if (file_exists($filePath)) {
            header('Content-Description: File Transfer');
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header('Content-Disposition: attachment; filename="Binuns_Stock_Report.xlsx"');
            header('Expires: 0');
            header('Cache-Control: must-revalidate');
            header('Pragma: public');
            header('Content-Length: ' . filesize($filePath));
            flush(); 
            readfile($filePath);
            exit;
        }
    }
    wp_die('File not found or access denied.');
}

add_action('admin_post_download_updated_spreadsheet', 'handle_download_updated_spreadsheet');
