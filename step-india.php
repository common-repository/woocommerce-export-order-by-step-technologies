<?php
/**
 * Plugin Name: Woocommerce export order by step-india.com
 * Description: Plugin will enable export of selected orders in CSV format
 * Author: S.T.E.P. Technologies
 * Author URI: http://step-india.com
 * License: GPL2
 * Version: 0.6
 * Requires at least: 3.3
 * Tested up to: 4.5.2
 * @package WooCommerce Export Order
 * @category Core

 */

if ( ! defined( 'ABSPATH' ) ) { 
    exit; // Exit if accessed directly
}
if ( !in_array( 'woocommerce/woocommerce.php', apply_filters( 'active_plugins', get_option( 'active_plugins' ) ) ) ) {
	//Woocommerce is not active, plugin can't serve anything
	return;
}

//export filter
 add_action( 'restrict_manage_posts', 'wceoExportOrders', 99 );
 function wceoExportOrders(){
        if (!isset($_GET['post_type']) || (isset($_GET['post_type']) && 'shop_order'!=$_GET['post_type'])){
             return;
        }
	
 	echo "<select name='shop_order_export' id='dropdown_shop_order_export'>
 			<option value='' >".__('All orders', 'wceo')."</option>
 			<option value='0' ".(isset($_GET['shop_order_export']) && ''!=$_GET['shop_order_export'] && 0==$_GET['shop_order_export'] ?'selected="selected"':'')." >".__('Not exported', 'wceo')."</option>
 			<option value='1' ".(isset($_GET['shop_order_export']) && 1==$_GET['shop_order_export'] ?'selected="selected"':'').">".__('All Exported', 'wceo')."</option>";
	$asPrevDates = get_option('wceoExportOrderDates');
	$sDTformat = get_option('date_format').' '.get_option('time_format');
	foreach($asPrevDates as $iDateTime){
		echo '<option value="'.$iDateTime.'" '.(isset($_GET['shop_order_export']) && $iDateTime==$_GET['shop_order_export'] ?'selected="selected"':'').'>Exported on '.date($sDTformat,$iDateTime).'</option>';
	}
	echo "
 	</select>
 			";
 			wc_enqueue_js( "

			jQuery('select#dropdown_shop_order_export').css('width', '200px');");
 }
 
 //bulk edit options
 add_action( 'admin_footer', 'wceoBulkExportOption' , 99 );
 function wceoBulkExportOption(){
 	global $post_type;
 	
 	if ( 'shop_order' == $post_type ) {
 		?>
 		<script type="text/javascript">
			jQuery(function() {
				jQuery('<option>').val('export_orders').text('<?php _e( 'Export orders', 'wceo' )?>').appendTo("select[name='action']");
				jQuery('<option>').val('export_orders').text('<?php _e( 'Export orders', 'wceo' )?>').appendTo("select[name='action2']");
			});
		</script>
 		<?php 
 	}
 }
 
 //perform actual export
 add_action( 'load-edit.php', 'wceoBulkExportData' );
 function wceoBulkExportData(){
 	$wp_list_table = _get_list_table( 'WP_Posts_List_Table' );
 	$action = $wp_list_table->current_action();
 	
 	$post_ids = array_map( 'absint', (array) $_REQUEST['post'] );
 	$changed = 0;
 	switch ( $action ) {
 		case 'export_orders':
 			$strMetaVal = time();
 			$strMetaKey = 'wceoIsDownloaded';
			
			header('Content-Disposition: attachment; filename="wceo'.date('YmdHis').'.xls"');
			echo '"order_id","order_number","date","status","shipping_total","shipping_tax_total","tax_total","cart_discount","order_discount","discount_total","order_total","payment_method","shipping_method","customer_id","billing_first_name","billing_last_name","billing_company","billing_email","billing_phone","billing_address_1","billing_address_2","billing_postcode","billing_city","billing_state","billing_country","shipping_first_name","shipping_last_name","shipping_address_1","shipping_address_2","shipping_postcode","shipping_city","shipping_state","shipping_country","shipping_company","customer_note","coupons","order_notes","line_items"';//,"shipping_items","tax_items"';
			foreach ( $post_ids as $post_id ) {
					
				wceoGetExportOrderLine($post_id);
				//Create/update two meta keys 1. For specific date and time and 2. for maintaining status of export
				update_post_meta($post_id, $strMetaKey.$strMetaVal, '1');
				update_post_meta($post_id, $strMetaKey, '1');
				$changed++;
			}
			//At least one order needs to be exported in order to build export sets
			if ($changed>0){
				$asPrevDates = get_option('wceoExportOrderDates');
				if (!is_array($asPrevDates)){
					$asPrevDates = array($asPrevDates);
				}
				$asPrevDates[] = $strMetaVal;
				update_option('wceoExportOrderDates', $asPrevDates);
			}
			die();
 			break;
 		
 		default:
 			return;
 	}
 	
 	
 	
 	
 	
 	
 	$sendback = add_query_arg( array( 'post_type' => 'shop_order', $report_action => true, 'changed' => $changed, 'ids' => join( ',', $post_ids ) ), '' );
 	wp_redirect( $sendback );
 	exit();
 	
 }
 
 function wceoGetExportOrderLine($intOrder){
 	$objOrder = new WC_Order( $intOrder );
 	$aryCustomerNotes = $objOrder->get_customer_order_notes();
 	$strCustomerNotes = '';
 	foreach($aryCustomerNotes as $valCustomerNote){
 		$strCustomerNotes .= strip_tags($valCustomerNote->comment_content).'|';
 	}
 	
 	$aryOrderNotes = wceoGetOrderNotes($objOrder);
 	$strOrderNotes = '';
 	foreach($aryOrderNotes as $valOrderNote){
 		$strOrderNotes .= strip_tags($valOrderNote->comment_content).'|';
 	}
 	
 	echo '
'.$intOrder.',"'.str_replace('"','""',$objOrder->get_order_number()).'","'.str_replace('"','""',$objOrder->order_date).'","'.str_replace('"','""',$objOrder->status).'",'.
	(method_exists($objOrder, 'get_shipping')?$objOrder->get_shipping():$objOrder->get_total_shipping()).','.$objOrder->get_shipping_tax().','.$objOrder->get_total_tax().','.$objOrder->get_cart_discount().','.$objOrder->get_order_discount().','.
 	$objOrder->get_total_discount().','.(method_exists($objOrder, 'get_order_total')?$objOrder->get_order_total():$objOrder->get_total()).',"'.str_replace('"','""',get_post_meta($intOrder,'_payment_method_title',true)).'","'.
 	str_replace('"','""',$objOrder->get_shipping_method()).'",'.$objOrder->customer_user.',"'.str_replace('"','""',$objOrder->billing_first_name).'","'.
 	$objOrder->billing_last_name.'","'.str_replace('"','""',$objOrder->billing_company).'","'.str_replace('"','""',$objOrder->billing_email).'","'.
 	str_replace('"','""','Contact Number').'","'.str_replace('"','""',$objOrder->billing_address_1).'","'.str_replace('"','""',$objOrder->billing_address_2).'","'.
 	str_replace('"','""',$objOrder->billing_postcode).'","'.str_replace('"','""',$objOrder->billing_city).'","'.str_replace('"','""',$objOrder->billing_state).'","'.
 	str_replace('"','""',$objOrder->billing_country).'","'.str_replace('"','""',$objOrder->shipping_first_name).'","'.str_replace('"','""',$objOrder->shipping_last_name).'","'.
 	str_replace('"','""',$objOrder->shipping_address_1).'","'.str_replace('"','""',$objOrder->shipping_address_2).'","'.str_replace('"','""',$objOrder->shipping_postcode).'","'.
 	str_replace('"','""',$objOrder->shipping_city).'","'.str_replace('"','""',$objOrder->shipping_state).'","'.str_replace('"','""',$objOrder->shipping_country).'","'.
 	str_replace('"','""',$objOrder->shipping_company).'","'.str_replace('"','""',$strCustomerNotes).'","'.str_replace('"','""',$objOrder->get_cart_discount_to_display()).'","'.
 	str_replace('"','""',$strOrderNotes).'",';	
 	 foreach($objOrder->get_items() as $keyOI=>$valOI){
 	 	echo '"name:'.str_replace('"','""',$valOI['name'].'|sku:'.(isset($valOI['sku'])?$valOI['sku']:'').'|quantity:'.$valOI['qty'].'|total:'.$valOI['line_total']).';"';
 	 }
 	 
 	 echo ',';//,Shipping_items
 	 
 	 //echo '';//Tax_items
 	 
 	 echo implode('|', $objOrder->get_used_coupons()).','.implode('|', $objOrder->get_customer_order_notes());
 	 
 	
 	
 }
 
 //Apply filters as per user selection
 add_action( 'pre_get_posts', 'wceoFilterOrdersExport' );
 function wceoFilterOrdersExport($query){
	if ( $query->query_vars['post_type'] == 'shop_order' && isset( $_GET['shop_order_export'] ) && ''!=$_GET['shop_order_export'] && $_GET['shop_order_export'] >= 0 ) {
		if ('0'==$_GET['shop_order_export'] ){
			$query->set('meta_query', array(
					'relation' => 'OR', 
					array('key' => 'wceoIsDownloaded', 
						'compare' => 'NOT EXISTS'),
					array('key' => 'wceoIsDownloaded',
							'value'=>'0')
						));
		}elseif ('1'==$_GET['shop_order_export']){
			$query->set('meta_query', array(
				array('key' => 'wceoIsDownloaded',
							'value'=>'1')
						));
			
		}else{
			$query->set('meta_query', array(
				array('key' => 'wceoIsDownloaded'.$_GET['shop_order_export'],
							'value'=>'1')
						));
		}
		
	}
	return $query;
 }
 
 //Get order notes
 function wceoGetOrderNotes($objOrder) {
 
 	$notes = array();
 
 	$args = array(
 			'post_id' => $objOrder->id,
 			'approve' => 'approve',
 			'type' => ''
 	);
 
 	remove_filter( 'comments_clauses', array( 'WC_Comments', 'exclude_order_comments' ) );
 
 	$comments = get_comments( $args );
 
 	foreach ( $comments as $comment ) {
 		$is_customer_note = get_comment_meta( $comment->comment_ID, 'is_customer_note', false );
 		$comment->comment_content = make_clickable( $comment->comment_content );
 		if ( $is_customer_note ) {
 			$notes[] = $comment;
 		}
 	}
 
 	add_filter( 'comments_clauses', array( 'WC_Comments', 'exclude_order_comments' ) );
 
 	return (array) $notes;
 
 }

?>
