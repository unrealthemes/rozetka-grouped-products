<?php
/*
 * Template Name: Rozetka
 */

ini_set('max_execution_time', '6000'); //10 часов
set_time_limit(0);

// get_header(); 
date_default_timezone_set('Europe/Kyiv');
$filename = __DIR__ . '/product-grouped.xls';
$status = get_option( 'procces_status' );
$date = get_option( 'procces_date' );
?>

<!DOCTYPE html>
<html lang="ru">
    <head>
        <META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=utf-8" />
        <TITLE>Rozetka</TITLE>
        <style>
            body {margin:0;padding:0;font: 12px Tahoma;}
            h1, h2 {font-size:20px;color:#1F84FF;margin-bottom:20px;margin-top:0;font-weight:normal;line-height:30px;}
            h2 {
                padding-left: 10px;
            }
            a {color:#1873b4;}
            body {
                background-color: #f5f5f5;
            }
            .header {
                display: flex;
                justify-content: space-between;
                background-color: #221f1f;
                padding: 10px;
            }
            .header span {
                background: #f5f5f5;
                border-radius: 15px;
                padding: 5px
            }
            .body {
                padding: 10px;
            }
            input[type="submit"] {
                cursor: pointer;
                background-color: #00a046;
                color: #fff;
                font-size: 16px;
                height: 40px;
                line-height: 40px;
                border: none;
                border-radius: 8px;
                box-sizing: border-box;
                display: inline-block;
                font-family: Rozetka,BlinkMacSystemFont,-apple-system,Arial,Segoe UI,Roboto,Helvetica,sans-serif;
                margin: 0;
                outline: none;
                padding-left: 16px;
                padding-right: 16px;
                position: relative;
                text-align: center;
                transition-duration: .2s;
                transition-property: color,background-color,border-color;
                transition-timing-function: ease-in-out;
            }
            .file-wrapper {
                padding: 10px;
            }
            .file-wrapper a {
                display: flex;
                align-items: center;
                justify-content: space-between;
                width: 170px;
            }
            .file-wrapper img {
                width: 50px;
            }
            .proccess-wrapper img {
                width: 50px;
            }
            .proccess-wrapper {
                padding: 10px 10px 30px 10px;
                display: flex;
                align-items: center;
                justify-content: space-between;
                width: 400px;
            }
        </style>
        
    </head>
    <body>
    
            <div class="header">
                <img alt="Rozetka Logo" src="https://content2.rozetka.com.ua/widget_logotype/full/original/229862237.svg">
                <span>
                    <img alt="Farbaua Logo" src="<?php echo get_template_directory_uri() . '/img/farbaua.png'; ?>">
                </span>
            </div>
            <div class="body">
                <h1>Группировка товаров</h1>
                <form action="" method="post" enctype="multipart/form-data">
                    <input type="file" name="products">
                    <input type="submit" name="start" value="Начать">
                </form>
            </div>

            <h2>Процесс</h2>
            <div class="proccess-wrapper">

                <?php if ($status) : ?>
                    <img src="<?php echo get_template_directory_uri() . '/img/preloader.gif'; ?>">
                <?php else : ?>
                    <img src="<?php echo get_template_directory_uri() . '/img/completed.png'; ?>">

                    <?php if ($date) : ?>
                        <span>
                            Дата последней генерации файла: 
                            <strong>
                                <?php echo $date; ?>
                            </strong>
                        </span>
                    <?php endif; ?>

                <?php endif; ?>

            </div>

            <?php if (file_exists($filename)) : ?>
                <h2>Файл</h2>
                <div class="file-wrapper">
                    <a href="<?php echo get_template_directory_uri(); ?>/product-grouped.xls" download>
                        <img src="<?php echo get_template_directory_uri(); ?>/img/excel.png" alt="Excel">
                        <span>product-grouped.xls</span>
                    </a>
                </div>
            <?php endif; ?>

    </body>
</html>

<?php 

if ( isset($_POST['start']) ) {
 
    require_once __DIR__ . '/lib/PHPExcel-1.8/Classes/PHPExcel.php';
    require_once __DIR__ . '/lib/PHPExcel-1.8/Classes/PHPExcel/Writer/Excel2007.php';
    require_once __DIR__ . '/lib/PHPExcel-1.8/Classes/PHPExcel/IOFactory.php';

    $file_name = basename($_FILES["products"]["name"]);
    $target_file_path = $file_name;
    $file_type = pathinfo($target_file_path, PATHINFO_EXTENSION);
    
    if ( $file_type != 'xls' ) {
        // $message = "Sorry, only CSV files are allowed to upload.";
        return false;
    }

    $products = [];
    $objPHPExcel = PHPExcel_IOFactory::load( $_FILES["products"]['tmp_name'] );
    //  Get worksheet dimensions
    $sheet = $objPHPExcel->getSheet(0); 
    $highestRow = $sheet->getHighestRow(); 
    $highestColumn = $sheet->getHighestColumn();
    //  Loop through each row of the worksheet in turn
    for ($row = 1; $row <= $highestRow; $row++){ 
        //  Read a row of data into an array
        $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, NULL, TRUE, FALSE);
        
        if ( $row == 1 ) {
            continue;
        }
        
        if ( empty($rowData[0][0]) ) {
            continue;
        }

        if ( ! strripos($rowData[0][2], 'краска') ) {
            continue;
        }
        
        if ( strripos($rowData[0][2], 'аэрозольный баллон') ) {
            continue;
        }

        $products[ $rowData[0][0] ] = $rowData[0][2];
    }
    
    update_option('products', $products);
    update_option( 'procces_status', true );

    set_shedule_grouped_products();

    ut_redirect($_SERVER['REQUEST_URI']);
}

// get_footer();