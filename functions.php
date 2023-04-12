<?php 

add_action( 'ut_grouped_products', 'ut_grouped_product_by_parts', 10, 1 );
add_filter( 'cron_schedules', 'ut_add_cron_recurrence_interval' );

function set_shedule_grouped_products() {
    // date_default_timezone_set('Asia/Tbilisi');
    date_default_timezone_set('UTC');
    $interval = 'every_15_minutes';
    $time = time();
    // remove shadule event for create new shedule with another interval
    wp_clear_scheduled_hook( 'ut_grouped_products' );
    wp_schedule_event( $time, $interval, 'ut_grouped_products' );
}



function ut_grouped_product_by_parts() {

    $products = get_option('products', true);
    $groups = get_option('groups', true);
    update_option( 'procces_status', true );

    if ( !$products || empty($products) ) {
        ut_create_excel($groups);
        wp_clear_scheduled_hook( 'ut_grouped_products' );
        update_option( 'procces_status', false );
        update_option( 'procces_date', wp_date("F j, Y, g:i") );
        update_option('products', []);
        update_option('groups', []);
        return false;
    }

    if ( empty($groups) || !is_array($groups) ) {
        $groups = [];
    }

    $count = 1;
    foreach ((array)$products as $rozetka_id => $product_name) {
        $product_name_g = substr($product_name, 0, -9);
        $i = true;

        if ( $count > 100 ) {
            break;
        }

        if (!empty($groups) && is_array($groups)) {
            foreach ((array)$groups as $key => $group) {
                foreach ((array)$group as $group_rozetka_id =>  $group_product_name) {
                    $group_product_name_g = substr($group_product_name, 0, -9);
                    $sim = similar_text($product_name_g, $group_product_name_g, $perc);

                    if ( $perc >= 99 ) {
                        $groups[ $key ][ $rozetka_id ] = $product_name;
                        unset($products[$rozetka_id]);
                        $i = false;
                        break 2;
                    } 
                }
            }

            if ($i) {
                $groups[][ $rozetka_id ] = $product_name;
                unset($products[$rozetka_id]);
            }

        } else {
            $groups[][ $rozetka_id ] = $product_name;
            unset($products[$rozetka_id]);
        }

        $count++;
    }

    update_option('products', $products);

    if ( $groups ) {
        update_option('groups', $groups);
    } else {
        update_option('groups', []);
    }
    
    set_shedule_grouped_products();

}


function ut_create_excel($groups) {

    require_once __DIR__ . '/lib/PHPExcel-1.8/Classes/PHPExcel.php';
    require_once __DIR__ . '/lib/PHPExcel-1.8/Classes/PHPExcel/Writer/Excel2007.php';
    require_once __DIR__ . '/lib/PHPExcel-1.8/Classes/PHPExcel/IOFactory.php';

    $xls = new PHPExcel();
    $xls->setActiveSheetIndex(0);
    $sheet = $xls->getActiveSheet();
    $sheet->setTitle('Групування товарів');
    $sheet->getColumnDimension("A")->setAutoSize(true); // ->setWidth(200);
    $sheet->getColumnDimension("B")->setAutoSize(true); // ->setWidth(200);
    $sheet->getColumnDimension("C")->setAutoSize(true); // ->setWidth(1000);
    $sheet->getColumnDimension("D")->setAutoSize(true); // ->setWidth(200);

    if ($groups) {
        $k = 1;
        foreach ($groups as $group) {

            $sheet->setCellValueExplicit( "A" . $k, 'Код товара в розетке', PHPExcel_Cell_DataType::TYPE_STRING );
            $sheet->setCellValueExplicit( "B" . $k, 'ID товара у продавца', PHPExcel_Cell_DataType::TYPE_STRING );
            $sheet->setCellValueExplicit( "C" . $k, 'Название товара', PHPExcel_Cell_DataType::TYPE_STRING );
            $sheet->setCellValueExplicit( "D" . $k, 'Сгруппировать по: ', PHPExcel_Cell_DataType::TYPE_STRING );
            //
            $sheet->getStyle("A" . $k)->getFont()->setBold(true);
            $sheet->getStyle("B" . $k)->getFont()->setBold(true);
            $sheet->getStyle("C" . $k)->getFont()->setBold(true);
            $sheet->getStyle("D" . $k)->getFont()->setBold(true);
            //
            $sheet->getStyle("A" . $k)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            $sheet->getStyle("B" . $k)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            $sheet->getStyle("C" . $k)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
            $sheet->getStyle("D" . $k)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

            $k++;

            foreach ($group as $group_rozetka_id => $group_product_name) {

                $sheet->setCellValueExplicit( "A" . $k, $group_rozetka_id, PHPExcel_Cell_DataType::TYPE_STRING );
                $sheet->getStyle("A" . $k)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
                $sheet->setCellValueExplicit( "B" . $k, '', PHPExcel_Cell_DataType::TYPE_STRING );
                $sheet->setCellValueExplicit( "C" . $k, $group_product_name, PHPExcel_Cell_DataType::TYPE_STRING );
                $sheet->setCellValueExplicit( "D" . $k, 'объему', PHPExcel_Cell_DataType::TYPE_STRING );

                $k++;
            }
        }
    }

    $objWriter = new PHPExcel_Writer_Excel5($xls);
    $objWriter->save(__DIR__ . '/product-grouped.xls');

}

function ut_add_cron_recurrence_interval( $schedules ) {

    $schedules['every_1_minute'] = [
        'interval'  => 60,
        'display'   => __( 'Every 1 Minute' )
    ];

    $schedules['every_15_minutes'] = [
        'interval'  => 900,
        'display'   => __( 'Every 15 Minutes' )
    ];

    $schedules['every_25_minutes'] = [
        'interval'  => 1500,
        'display'   => __( 'Every 25 Minutes' )
    ];
    
    $schedules['1_hour'] = [
        'interval'  => 3600,
        'display'   => __( '1 Hour' )
    ];

    $schedules['2_hours'] = [
        'interval'  => 7200,
        'display'   => __( '2 Hours' )
    ];
     
    return $schedules;
}

function ut_redirect($url) {

    $string = '<script type="text/javascript">';
    $string .= 'window.location = "' . $url . '"';
    $string .= '</script>';

    echo $string;
}