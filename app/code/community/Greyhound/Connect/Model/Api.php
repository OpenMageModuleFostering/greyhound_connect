<?php
/*
 * The MIT License (MIT)
 *
 * Copyright (c) 2013-2016 digital guru GmbH & Co. KG
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 */

/**
 * GREYHOUND Connect extension custom api.
 *
 * @category Greyhound
 * @package Greyhound_Connect
 * @author digital guru GmbH &amp; Co. KG <develop@greyhound-software.com>
 * @copyright 2013-2016 digital guru GmbH &amp; Co. KG
 * @license http://opensource.org/licenses/MIT MIT License
 * @link greyhound-software.com
 */
class Greyhound_Connect_Model_Api extends Mage_Api_Model_Resource_Abstract
{
  /**
   * Cached payment method titles (by payment method code).
   * @var array
   */
  protected $_paymentMethods = array();

  /**
   * Timezones of the stores (by store_id).
   * @var array
   */
  protected $_timezones = array();


  /**
   * Returns basic module and shop information.
   *
   * @return object
   */
  public function info()
  {
    $moduleVersion = '';

    /* @var Varien_Simplexml_Element $moduleConfig */
    $moduleConfig = Mage::getConfig()->getModuleConfig('Greyhound_Connect');

    if($moduleConfig)
    {
      $moduleConfig = $moduleConfig->asArray();

      if(isset($moduleConfig['version']))
        $moduleVersion = $moduleConfig['version'];
    }

    return (object)array(
      'module_version' => $moduleVersion,
      'shop_version' => Mage::getVersion(),
      'shop_edition' => function_exists('Mage::getEdition') ? Mage::getEdition() . ' Edition' : ''
    );
  }

  /**
   * Returns orders of customers matching the filter criteria.
   *
   * @param array $filters filters
   * @param int   $limit   maximum number of results (0 = unlimited)
   * @return array
   */
  public function items($filters, $limit = 0)
  {
    // Fetch orders matching the filter:
    $orders = $this->_getOrderData($filters, $limit);

    foreach($orders as $orderId => $order)
    {
      $order['billing_address'] = (object)$order['billing_address'];
      $orders[$orderId] = $order;
    }

    if(empty($orders))
      return array();

    // Add order comments to these orders:
    foreach($this->_getOrderComments($orders) as $orderId => $comments)
      if(isset($orders[$orderId]))
        $orders[$orderId]['comments'] = $comments;

    // Add payment data to these orders:
    foreach($this->_getPayments($orders) as $orderId => $payments)
      if(isset($orders[$orderId]))
        $orders[$orderId]['payments'] = $payments;

    // Add invoice data to these orders:
    foreach($this->_getInvoices($orders) as $orderId => $invoices)
      if(isset($orders[$orderId]))
        $orders[$orderId]['invoices'] = $invoices;

    // Add credit memo data to these orders:
    foreach($this->_getCreditmemos($orders) as $orderId => $creditmemos)
      if(isset($orders[$orderId]))
        $orders[$orderId]['creditmemos'] = $creditmemos;

    // Add shipment data to these orders:
    foreach($this->_getShipments($orders) as $orderId => $shipments)
      if(isset($orders[$orderId]))
        $orders[$orderId]['shipments'] = $shipments;

    // Add order item data to these orders:
    foreach($this->_getOrderItems($orders) as $orderId => $items)
      if(isset($orders[$orderId]))
        $orders[$orderId]['items'] = $items;

    // Return the orders:
    $result = array();

    foreach($orders as $order)
      $result[] = array('json' => json_encode($order));

    return $result;
  }

  /**
   * Returns a list of filters and tables for the orders collection query.
   * The array contains the filters in the key 'filters' and the tables in the key 'tables'.
   *
   * @param Mage_Sales_Model_Resource_Order_Collection $collection order query collection
   * @param array                                      $filters    request filters
   */
  protected function _applyFiltersToOrderQuery($collection, $filters)
  {
    // We will add each filter to the query and join tables as necessary.
    // If the filter value is an array, then an 'IN' query condition will be used, otherwise either an 'eq' or a 'like'
    // condition will be used, depending on whether the filter field is expected to support 'fuzzy' search terms.

    if(!is_array($filters))
      $this->_fault('filters_invalid', 'No filters specified');

    $tables = array();

    foreach($filters as $field => $value)
    {
      if(is_array($value))
        $operator = 'in';
      else
        $operator = 'eq';

      switch($field)
      {
        // Address data:
        case 'company':
        case 'firstname':
        case 'lastname':
        case 'street':
        case 'postcode':
        case 'city':
          if($operator == 'eq')
          {
            $operator = 'like';
            $value .= '%';
          }

          // Billing address:
          if(!isset($tables['billtbl']))
          {
            $tables['billtbl'] = true;
            $collection->getSelect()->joinLeft(array('billtbl' => $collection->getTable('sales/order_address')), 'main_table.billing_address_id = billtbl.entity_id', array());
          }

          $filter = array('attribute' => 'billtbl.' . $field, $operator => $value);
          $collection->addFieldToSearchFilter($filter['attribute'], $filter);

          // Shipping address:
          if(!isset($tables['shiptbl']))
          {
            $tables['shiptbl'] = true;
            $collection->getSelect()->joinLeft(array('shiptbl' => $collection->getTable('sales/order_address')), 'main_table.shipping_address_id = shiptbl.entity_id', array());
          }

          $filter = array('attribute' => 'shiptbl.' . $field, $operator => $value);
          $collection->addFieldToSearchFilter($filter['attribute'], $filter);
          break;

        // Order number:
        case 'order_id':
          $filter = array('attribute' => 'main_table.increment_id', $operator => $value);
          $collection->addFieldToSearchFilter($filter['attribute'], $filter);
          break;

        // Invoice number:
        case 'invoice_id':
          if(!isset($tables['invtbl']))
          {
            $tables['invtbl'] = true;
            $collection->getSelect()->joinLeft(array('invtbl' => $collection->getTable('sales/invoice')), 'main_table.entity_id = invtbl.order_id', array());
          }

          $filter = array('attribute' => 'invtbl.increment_id', $operator => $value);
          $collection->addFieldToSearchFilter($filter['attribute'], $filter);
          break;

        // Credit memo number:
        case 'creditmemo_id':
          if(!isset($tables['crmemtbl']))
          {
            $tables['crmemtbl'] = true;
            $collection->getSelect()->joinLeft(array('crmemtbl' => $collection->getTable('sales/creditmemo')), 'main_table.entity_id = crmemtbl.order_id', array());
          }

          $filter = array('attribute' => 'crmemtbl.increment_id', $operator => $value);
          $collection->addFieldToSearchFilter($filter['attribute'], $filter);
          break;

        // Other order data:
        default:
          $filter = array('attribute' => 'main_table.' . $field, $operator => $value);
          $collection->addFieldToSearchFilter($filter['attribute'], $filter);
          break;
      }
    }
  }

  /**
   * Returns basic order data for a list of customers.
   *
   * @param object $filters filter object
   * @param int   $limit    maximum number of results (0 = unlimited)
   * @return array
   */
  protected function _getOrderData($filters, $limit)
  {
    $collection = Mage::getModel('sales/order')->getCollection();
    $this->_applyFiltersToOrderQuery($collection, $filters);
    $collection->addAttributeToSort('created_at', 'desc');

    if($limit > 0)
      $collection->setPageSize($limit);

    $orderFields = array('entity_id', 'increment_id', 'ext_order_id', 'state', 'status', 'created_at', 'updated_at', 'store_id', 'store_name', 'customer_id', 'ext_customer_id', 'order_currency_code', 'grand_total', 'total_paid', 'discount_amount', 'discount_description', 'shipping_incl_tax', 'customer_dob', 'shipping_method', 'shipping_description');

    foreach($orderFields as $field)
      $collection->addAttributeToSelect($field);

    $addressFields = array('company', 'firstname', 'lastname', 'street', 'postcode', 'city', 'region', 'country_id', 'telephone', 'fax', 'email');
    $addressTable = $collection->getTable('sales/order_address');

    // Billing address:
    $billingAlias = 'billing_o_a';
    $billingQueryFields = array();

    foreach($addressFields as $field)
    {
      $collection->addFilterToMap('billaddr_' . $field, $billingAlias . '.' . $field);
      $collection->addExpressionFieldToSelect('billaddr_' . $field, '{{billaddr_' . $field . '}}', array('billaddr_' . $field => $billingAlias . '.' . $field));
      $billingQueryFields[] = $billingAlias . '.' . $field;
    }

    $collection->getSelect()->joinLeft(array($billingAlias => $addressTable), 'main_table.billing_address_id = ' . $billingAlias . '.entity_id', $billingQueryFields);

    // Shipping address:
    $shippingAlias = 'shipping_o_a';
    $shippingQueryFields = array();

    foreach($addressFields as $field)
    {
      $collection->addFilterToMap('shipaddr_' . $field, $shippingAlias . '.' . $field);
      $collection->addExpressionFieldToSelect('shipaddr_' . $field, '{{shipaddr_' . $field . '}}', array('shipaddr_' . $field => $shippingAlias . '.' . $field));
      $shippingQueryFields[] = $shippingAlias . '.' . $field;
    }

    $collection->getSelect()->joinLeft(array($shippingAlias => $addressTable), 'main_table.shipping_address_id = ' . $shippingAlias . '.entity_id', $shippingQueryFields)->group('main_table.entity_id');


    // Get order data:
    $orders = array();

    foreach($collection as $entry)
    {
      /* @var Mage_Sales_Model_Order $entry */
      $entry = $entry->getData();
      $order = array();

      foreach($orderFields as $field)
        $order[$field] = $entry[$field];

      $order['created_at'] = $this->_convertTime($order['created_at'], $order['store_id']);
      $order['updated_at'] = $this->_convertTime($order['updated_at'], $order['store_id']);

      $billingAddress = array();
      $shippingAddress = array();

      foreach($addressFields as $field)
      {
        $billingAddress[$field] = $entry['billaddr_' . $field];
        $shippingAddress[$field] = $entry['shipaddr_' . $field];
      }

      $order['billing_address'] = $billingAddress;
      $order['shipping_address'] = $shippingAddress;
      $order['payments'] = array();
      $order['invoices'] = array();
      $order['creditmemos'] = array();
      $order['shipments'] = array();
      $order['comments'] = array();
      $order['items'] = array();

      $adminHelper = Mage::helper('adminhtml');/* @var Mage_Adminhtml_Helper_Data $adminHelper */
      $order['url'] = preg_replace('#SID=[a-z0-9]*#', '', preg_replace('#/key/[^/]*#', '', $adminHelper->getUrl('adminhtml/sales_order/view', array('order_id' => $order['entity_id']))));
      $order['customer_url'] = $adminHelper->getUrl('adminhtml/customer/edit', array('id' => $order['customer_id']));

      $orders[$order['entity_id']] = $order;
    }

    return $orders;
  }

  /**
   * Returns order comments for a list of orders.
   *
   * @param array $orders orders
   * @return array
   */
  protected function _getOrderComments($orders)
  {
    $collection = Mage::getModel('sales/order_status_history')->getCollection();
    $collection->addFieldToFilter('parent_id', array_keys($orders));
    $collection->addAttributeToSelect('parent_id');
    $collection->addAttributeToSelect('comment');
    $collection->addAttributeToSelect('created_at');

    $result = array();

    // Group payments by order id:
    foreach($collection as $entry)
    {
      /* @var Mage_Sales_Model_Order_Status_History $entry */
      $entry = $entry->getData();
      $orderId = $entry['parent_id'];

      if(!isset($orders[$orderId]) || strlen($entry['comment']) == 0)
        continue;

      $storeId = $orders[$orderId]['store_id'];

      $comment = array(
        'created_at' => $this->_convertTime($entry['created_at'], $storeId),
        'comment' => $entry['comment']
      );

      if(!isset($result[$orderId]))
        $result[$orderId] = array($comment);
      else
        $result[$orderId][] = $comment;
    }

    return $result;
  }

  /**
   * Returns payment data for a list of orders.
   *
   * @param array $orders orders
   * @return array
   */
  protected function _getPayments($orders)
  {
    $collection = Mage::getModel('sales/order_payment')->getCollection();
    $collection->addFieldToFilter('parent_id', array_keys($orders));
    $collection->addAttributeToSelect('parent_id');
    $collection->addAttributeToSelect('method');

    $result = array();

    // Group payments by order id:
    foreach($collection as $entry)
    {
      /* @var Mage_Sales_Model_Order_Payment $entry */
      $entry = $entry->getData();
      $orderId = $entry['parent_id'];

      $paymentMethod = $entry['method'];

      $payment = array(
        'method' => $paymentMethod,
        'title' => $this->_getPaymentDescription($paymentMethod)
      );

      if(!isset($result[$orderId]))
        $result[$orderId] = array($payment);
      else
        $result[$orderId][] = $payment;
    }

    return $result;
  }

  /**
   * Returns invoice data for a list of orders.
   *
   * @param array $orders orders
   * @return array
   */
  protected function _getInvoices($orders)
  {
    $collection = Mage::getModel('sales/order_invoice')->getCollection();
    $collection->addFieldToFilter('order_id', array_keys($orders));
    $fields = array(
      'order_id',
      'entity_id',
      'increment_id',
      'created_at',
      'updated_at',
      'grand_total',
      'subtotal_incl_tax',
      'shipping_incl_tax',
      'discount_amount',
      'total_qty',
      'state'
    );

    foreach($fields as $field)
      $collection->addAttributeToSelect($field);

    $invoices = array();

    foreach($collection as $entry)
    {
      /* @var Mage_Sales_Model_Order_Invoice $entry */
      $entry = $entry->getData();
      $orderId = $entry['order_id'];

      if(!isset($orders[$orderId]))
        continue;

      $storeId = $orders[$orderId]['store_id'];
      $invoice = array();

      foreach($fields as $field)
        $invoice[$field] = $entry[$field];

      $invoice['store_id'] = $storeId;

      $invoice['created_at'] = $this->_convertTime($invoice['created_at'], $storeId);
      $invoice['updated_at'] = $this->_convertTime($invoice['updated_at'], $storeId);
      $invoice['comments'] = array();
      $invoice['items'] = array();

      $invoices[$invoice['entity_id']] = $invoice;
    }

    if(empty($invoices))
      return array();

    // Invoice comments:
    $collection = Mage::getModel('sales/order_invoice_comment')->getCollection();
    $collection->addFieldToFilter('parent_id', array_keys($invoices));
    $collection->addAttributeToSelect('parent_id');
    $collection->addAttributeToSelect('created_at');
    $collection->addAttributeToSelect('comment');

    foreach($collection as $entry)
    {
      /* @var Mage_Sales_Model_Order_Invoice_Comment $entry */
      $entry = $entry->getData();
      $parentId = $entry['parent_id'];

      if(!isset($invoices[$parentId]))
        continue;

      $storeId = $invoices[$parentId]['store_id'];

      $invoices[$parentId]['comments'][] = array(
        'created_at' => $this->_convertTime($entry['created_at'], $storeId),
        'comment' => $entry['comment']
      );
    }

    // Invoice items:
    $collection = Mage::getModel('sales/order_invoice_item')->getCollection();
    $collection->addFieldToFilter('parent_id', array_keys($invoices));
    $fields = array(
      'parent_id',
      'discount_amount',
      'price_incl_tax',
      'row_total_incl_tax',
      'qty',
      'product_id',
      'order_item_id',
      'sku',
      'name'
    );

    foreach($fields as $field)
      $collection->addAttributeToSelect($field);

    foreach($collection as $entry)
    {
      /* @var Mage_Sales_Model_Order_Invoice_Item $entry */
      $entry = $entry->getData();

      $parentId = $entry['parent_id'];

      if(!isset($invoices[$parentId]))
        continue;

      $invoiceItem = array();

      foreach($fields as $field)
        $invoiceItem[$field] = $entry[$field];

      $invoices[$parentId]['items'][] = $invoiceItem;
    }

    // Group invoices by order id:
    $result = array();

    foreach($invoices as $invoice)
    {
      $orderId = $invoice['order_id'];
      unset($invoice['order_id']);
      unset($invoice['store_id']);

      if(!isset($result[$orderId]))
        $result[$orderId] = array($invoice);
      else
        $result[$orderId][] = $invoice;
    }

    return $result;
  }

  /**
   * Returns credit memo data for a list of orders.
   *
   * @param array $orders orders
   * @return array
   */
  protected function _getCreditmemos($orders)
  {
    $collection = Mage::getModel('sales/order_creditmemo')->getCollection();
    $collection->addFieldToFilter('order_id', array_keys($orders));
    $fields = array(
      'order_id',
      'entity_id',
      'increment_id',
      'created_at',
      'updated_at',
      'grand_total',
      'adjustment_positive',
      'adjustment_negative',
      'subtotal_incl_tax',
      'shipping_incl_tax',
      'creditmemo_status',
      'state',
    );

    foreach($fields as $field)
      $collection->addAttributeToSelect($field);

    $creditmemos = array();

    foreach($collection as $entry)
    {
      /* @var Mage_Sales_Model_Order_Creditmemo $entry */
      $entry = $entry->getData();
      $orderId = $entry['order_id'];

      if(!isset($orders[$orderId]))
        continue;

      $storeId = $orders[$orderId]['store_id'];

      $creditmemo = array();

      foreach($fields as $field)
        $creditmemo[$field] = $entry[$field];

      $creditmemo['created_at'] = $this->_convertTime($creditmemo['created_at'], $storeId);
      $creditmemo['updated_at'] = $this->_convertTime($creditmemo['updated_at'], $storeId);
      $creditmemo['store_id'] = $storeId;
      $creditmemo['comments'] = array();
      $creditmemo['items'] = array();

      $creditmemos[$entry['entity_id']] = $creditmemo;
    }

    if(empty($creditmemos))
      return array();

    // Creditmemo comments:
    $collection = Mage::getModel('sales/order_creditmemo_comment')->getCollection();
    $collection->addFieldToFilter('parent_id', array_keys($creditmemos));
    $collection->addAttributeToSelect('parent_id');
    $collection->addAttributeToSelect('created_at');
    $collection->addAttributeToSelect('comment');

    foreach($collection as $entry)
    {
      /* @var Mage_Sales_Model_Order_Creditmemo_Comment $entry */
      $entry = $entry->getData();
      $parentId = $entry['parent_id'];

      if(!isset($creditmemos[$parentId]))
        continue;

      $storeId = $creditmemos[$parentId]['store_id'];

      $creditmemos[$parentId]['comments'][] = array(
        'created_at' => $this->_convertTime($entry['created_at'], $storeId),
        'comment' => $entry['comment']
      );
    }

    // Creditmemo items:
    $collection = Mage::getModel('sales/order_creditmemo_item')->getCollection();
    $collection->addFieldToFilter('parent_id', array_keys($creditmemos));
    $fields = array(
      'parent_id',
      'discount_amount',
      'price_incl_tax',
      'row_total_incl_tax',
      'qty',
      'product_id',
      'order_item_id',
      'sku',
      'name'
    );

    foreach($fields as $field)
      $collection->addAttributeToSelect($field);

    foreach($collection as $entry)
    {
      /* @var Mage_Sales_Model_Order_Creditmemo_Item $entry */
      $entry = $entry->getData();
      $parentId = $entry['parent_id'];

      if(!isset($creditmemos[$parentId]))
        continue;

      $creditmemoItem = array();

      foreach($fields as $field)
        $creditmemoItem[$field] = $entry[$field];

      $creditmemos[$parentId]['items'][] = $creditmemoItem;
    }

    // Group credit memos by order id:
    $result = array();

    foreach($creditmemos as $creditmemo)
    {
      $orderId = $creditmemo['order_id'];
      unset($creditmemo['order_id']);
      unset($creditmemo['store_id']);

      if(!isset($result[$orderId]))
        $result[$orderId] = array($creditmemo);
      else
        $result[$orderId][] = $creditmemo;
    }

    return $result;
  }

  /**
   * Returns shipment data for a list of orders.
   *
   * @param array $orders orders
   * @return array
   */
  protected function _getShipments($orders)
  {
    $collection = Mage::getModel('sales/order_shipment')->getCollection();
    $collection->addFieldToFilter('order_id', array_keys($orders));
    $fields = array(
      'entity_id',
      'order_id',
      'increment_id',
      'created_at',
      'updated_at'
    );

    foreach($fields as $field)
      $collection->addAttributeToSelect($field);

    $shipments = array();

    foreach($collection as $entry)
    {
      /* @var Mage_Sales_Model_Order_Shipment $entry */
      $entry = $entry->getData();
      $orderId = $entry['order_id'];

      if(!isset($orders[$orderId]))
        continue;

      $storeId = $orders[$orderId]['store_id'];
      $shipment = array();

      foreach($fields as $field)
        $shipment[$field] = $entry[$field];

      $shipment['created_at'] = $this->_convertTime($shipment['created_at'], $storeId);
      $shipment['updated_at'] = $this->_convertTime($shipment['updated_at'], $storeId);
      $shipment['store_id'] = $storeId;
      $shipment['comments'] = array();
      $shipment['tracking'] = array();

      $shipments[$entry['entity_id']] = $shipment;
    }

    if(empty($shipments))
      return array();

    $shipmentIds = array_keys($shipments);

    // Shipment comments:
    $collection = Mage::getModel('sales/order_shipment_comment')->getCollection();
    $collection->addFieldToFilter('parent_id', $shipmentIds);
    $collection->addAttributeToSelect('parent_id');
    $collection->addAttributeToSelect('created_at');
    $collection->addAttributeToSelect('comment');

    foreach($collection as $entry)
    {
      /* @var Mage_Sales_Model_Order_Shipment_Comment $entry */
      $entry = $entry->getData();
      $parentId = $entry['parent_id'];

      if(!isset($shipments[$parentId]))
        continue;

      $storeId = $shipments[$parentId]['store_id'];

      $shipments[$parentId]['comments'][] = array(
        'created_at' => $this->_convertTime($entry['created_at'], $storeId),
        'comment' => $entry['comment']
      );
    }

    // Shipment tracking:
    $collection = Mage::getModel('sales/order_shipment_track')->getCollection();
    $collection->addFieldToFilter('parent_id', $shipmentIds);
    $fields = array(
      'parent_id',
      'created_at',
      'updated_at',
      'track_number',
      'carrier_code',
      'title'
    );

    foreach($fields as $field)
      $collection->addAttributeToSelect($field);

    foreach($collection as $entry)
    {
      /* @var Mage_Sales_Model_Order_Shipment_Track $entry */
      $entry = $entry->getData();
      $parentId = $entry['parent_id'];

      if(!isset($shipments[$parentId]))
        continue;

      $storeId = $shipments[$parentId]['store_id'];

      $tracking = array();

      foreach($fields as $field)
        $tracking[$field] = $entry[$field];

      $tracking['created_at'] = $this->_convertTime($tracking['created_at'], $storeId);
      $tracking['updated_at'] = $this->_convertTime($tracking['updated_at'], $storeId);

      $shipments[$parentId]['tracking'][] = $tracking;
    }

    // Shipment items:
    $collection = Mage::getModel('sales/order_shipment_item')->getCollection();
    $collection->addFieldToFilter('parent_id', array_keys($shipments));
    $fields = array(
      'parent_id',
      'price',
      'row_total',
      'qty',
      'product_id',
      'order_item_id',
      'sku',
      'name'
    );

    foreach($fields as $field)
      $collection->addAttributeToSelect($field);

    foreach($collection as $entry)
    {
      /* @var Mage_Sales_Model_Order_Shipment_Item $entry */
      $entry = $entry->getData();
      $parentId = $entry['parent_id'];

      if(!isset($shipments[$parentId]))
        continue;

      $shipmentItem = array();

      foreach($fields as $field)
        $shipmentItem[$field] = $entry[$field];

      $shipments[$parentId]['items'][] = $shipmentItem;
    }

    // Group shipments by order id:
    $result = array();

    foreach($shipments as $shipment)
    {
      $orderId = $shipment['order_id'];
      unset($shipment['order_id']);
      unset($shipment['store_id']);

      if(!isset($result[$orderId]))
        $result[$orderId] = array($shipment);
      else
        $result[$orderId][] = $shipment;
    }

    return $result;
  }

  /**
   * Returns order item data for a list of orders.
   *
   * @param array $orders orders
   * @return array
   */
  protected function _getOrderItems($orders)
  {
    $collection = Mage::getModel('sales/order_item')->getCollection();
    $collection->addFieldToFilter('order_id', array_keys($orders));
    $fields = array(
      'order_id',
      'item_id',
      'parent_item_id',
      'store_id',
      'created_at',
      'updated_at',
      'product_type',
      'product_options',
      'sku',
      'name',
      'description',
      'qty_ordered',
      'qty_canceled',
      'qty_refunded',
      'qty_invoiced',
      'qty_shipped',
      'price_incl_tax',
      'row_total_incl_tax'
    );

    foreach($fields as $field)
      $collection->addAttributeToSelect($field);

    // First create a flat array of order items by id:
    $items = array();

    foreach($collection as $entry)
    {
      /* @var Mage_Sales_Model_Order_Item $entry */
      $entry = $entry->getData();

      if(!isset($orders[$entry['order_id']]))
        continue;

      $item = array();

      foreach($fields as $field)
        $item[$field] = $entry[$field];

      $item['items'] = array();
      $item['created_at'] = $this->_convertTime($item['created_at'], $item['store_id']);
      $item['updated_at'] = $this->_convertTime($item['updated_at'], $item['store_id']);
      unset($item['store_id']);

      $items[$item['item_id']] = $item;
    }

    // Then assign each sub-item to its parent:
    foreach($items as $item)
    {
      $parentId = $item['parent_item_id'];

      if($parentId && isset($items[$parentId]))
      {
        unset($item['order_id']); // The order_id is only needed for the root level items
        $items[$parentId]['items'][] = $item;
      }
    }

    // Group all root items by order id (the sub-items are contained within the root items):
    $result = array();

    foreach($items as $item)
    {
      if(!$item['parent_item_id'])
      {
        $orderId = $item['order_id'];
        unset($item['order_id']);

        if(!isset($result[$orderId]))
          $result[$orderId] = array($item);
        else
          $result[$orderId][] = $item;
      }
    }

    return $result;
  }

  /**
   * Returns the title of a payment method.
   *
   * @param string $paymentMethod payment method code
   * @return string
   */
  protected function _getPaymentDescription($paymentMethod)
  {
    if(!isset($this->_paymentMethods[$paymentMethod]))
    {
      $this->_paymentMethods[$paymentMethod] = '';

      $paymentHelper = Mage::helper('payment');/* @var Mage_Payment_Helper_Data $paymentHelper */
      $payment = $paymentHelper->getMethodInstance($paymentMethod);/* @var Mage_Payment_Model_Method_Abstract $payment */

      if($payment)
        $this->_paymentMethods[$paymentMethod] = $payment->getTitle();
    }

    return $this->_paymentMethods[$paymentMethod];
  }

  /**
   * Converts a date/time into the store's timezone.
   *
   * @param string $time    date/time string
   * @param int    $storeId store id
   * @return string
   */
  protected function _convertTime($time, $storeId)
  {
    if(!isset($this->_timezones[$storeId]))
    {
      $store = Mage::app()->getStore($storeId);

      if($store)
        $this->_timezones[$storeId] = $store->getConfig(Mage_Core_Model_Locale::XML_PATH_DEFAULT_TIMEZONE);
      else
        $this->_timezones[$storeId] = 'Europe/Berlin';
    }

    $timezone = $this->_timezones[$storeId];
    $date = new Zend_Date($time, 'YYYY-MM-dd HH:mm:ss');
    $date->setTimezone($timezone);

    return $date->get('YYYY-MM-dd HH:mm:ss');
  }
}