function parseOrders() {
  function getOrderDetails(orderId: string) {
    const orderDetailsLink =
      'https://www.amazon.com/gp/your-account/order-details?ie=UTF8&orderID=';

    return fetch(orderDetailsLink + orderId);
  }

  function getInvoiceDetails(orderId: string) {
    const invoiceDetailsLink =
      'https://www.amazon.com/gp/css/summary/print.html/ref=ppx_od_dt_b_invoice?ie=UTF8&orderID=';

    return fetch(invoiceDetailsLink + orderId);
  }

  function getTrackingInfo(trackingLink: string) {
    const trackingDetailsLink = `https://www.amazon.com${trackingLink}`;

    return fetch(trackingDetailsLink);
  }

  const orderIds = Array.from(
    document.querySelectorAll('.yohtmlc-order-id span:nth-child(2)')
  ).map((e) => (e as HTMLElement).innerText);

  const headers = [
    'Website',
    'Order ID',
    'Order Date',
    'Purchase Order Number',
    'Currency',
    'Unit Price',
    'Unit Price Tax',
    'Shipping Charge',
    'Total Discounts',
    'Total Owed',
    'Shipment Item Subtotal',
    'Shipment Item Subtotal Tax',
    'ASIN',
    'Product Condition',
    'Quantity',
    'Payment Instrument Type',
    'Order Status',
    'Shipment Status',
    'Ship Date',
    'Shipping Option',
    'Shipping Address',
    'Billing Address',
    'Carrier Name & Tracking Number',
    'Product Name',
    'Gift Message',
    'Gift Sender Name',
    'Gift Recipient Contact Details',
    'Vendor'
  ] as const;

  type Header = (typeof headers)[number];
  let csvString = headers.join(',') + '\n';

  orderIds.forEach((orderId) => {
    getOrderDetails(orderId).then((response) => {
      response.text().then((responseBody) => {
        const doc = new DOMParser().parseFromString(responseBody, 'text/html');

        const productTitleElement = doc.querySelectorAll(
          '.shipment .a-link-normal'
        )[1];
        const productTitle =
          (productTitleElement as HTMLElement)?.innerText || '';

        // trim whitespace and newlines from ends
        const productTitleTrimmed = productTitle.replace(/^\s+|\s+$/g, '');

        const productLink = productTitleElement?.getAttribute('href') || '';

        const trackingLink =
          doc.querySelector('.track-package-button a')?.getAttribute('href') ||
          '';

        const asin = productLink?.split('/')[3];

        const subTotals = (doc.querySelector('#od-subtotals') as HTMLElement)
          .innerText;

        let totalOwed = '';
        if (subTotals) {
          const matches = subTotals.match(/Grand Total:\s*(\$|USD)\s*(.*)/);
          if (matches && matches[2]) {
            totalOwed = matches[2];
          }
        }

        const orderDateMatches = (
          doc.querySelector('.order-date-invoice-item') as HTMLElement
        ).innerText.match('Ordered on (.*)');

        const orderDateString = orderDateMatches ? orderDateMatches[1] : '';
        const orderDate = new Date(orderDateString);
        orderDate.setHours(10);

        const orderDateISO = orderDate.toISOString();

        const isEasyBookPrep = (
          doc.querySelector(
            '.od-shipping-address-container .displayAddressDiv'
          ) as HTMLElement
        ).innerText.includes('POLARIS');

        const orderDetails: Record<Header, any> = {
          Website: 'zzAmazon.com',
          'Order ID': orderId,
          'Order Date': orderDateISO,
          'Purchase Order Number': 'zzNot Applicable',
          Currency: 'USD',
          'Unit Price': totalOwed,
          'Unit Price Tax': 'zztax',
          'Shipping Charge': 'zz0',
          'Total Discounts': 'zz0',
          'Total Owed': totalOwed,
          'Shipment Item Subtotal': totalOwed,
          'Shipment Item Subtotal Tax': 'zztax',
          ASIN: asin,
          'Product Condition': 'condition',
          Quantity: '1',
          'Payment Instrument Type': 'zzVisa',
          'Order Status': 'zzFor sale',
          'Shipment Status': 'zzShipped',
          'Ship Date': 'zzshipdate',
          'Shipping Option': 'zzStandard',
          'Shipping Address': isEasyBookPrep ? 'POLARIS' : 'Little Owl',
          'Billing Address': 'zzaddress',
          'Carrier Name & Tracking Number': 'zz',
          'Product Name': `"${productTitleTrimmed}"`,
          'Gift Message': 'zz',
          'Gift Sender Name': 'zz',
          'Gift Recipient Contact Details': 'zz',
          Vendor: 'vendor'
        };

        getInvoiceDetails(orderId).then((invResponse) => {
          invResponse.text().then((invResponseBody) => {
            const conditionMatches =
              invResponseBody.match(/Condition:\s*(.*)</);
            const condition = conditionMatches ? conditionMatches[1] : '';

            const vendorMatches = invResponseBody.match(/Sold by:\s*(.*)\s*\(/);
            const vendor = vendorMatches ? vendorMatches[1] : '';

            orderDetails['Product Condition'] = condition;
            orderDetails['Vendor'] = `"${vendor}"`;

            getTrackingInfo(trackingLink).then((trackingInfoResponse) => {
              trackingInfoResponse.text().then((trackingInfoResponseBody) => {
                const trackingInfoDoc = new DOMParser().parseFromString(
                  trackingInfoResponseBody,
                  'text/html'
                );

                const trackingNumber =
                  (
                    trackingInfoDoc.querySelector(
                      '.pt-delivery-card-trackingId'
                    ) as HTMLElement
                  )?.innerText ?? '';

                const trackingNumberOnly =
                  trackingNumber.match(/Tracking ID: (.*)/);

                orderDetails['Carrier Name & Tracking Number'] =
                  trackingNumberOnly?.[1] ?? '';

                // print as csv
                csvString +=
                  headers.map((key) => orderDetails[key as Header]).join(',') +
                  '\n';
                // console.log(
                //   headers.map((key) => orderDetails[key as Header]).join(',') + '\n'
                // );

                console.log(csvString);
              });
            });
          });
        });
      });
    });
  });
}

chrome.action.onClicked.addListener((tab) => {
  if (!tab || !tab.url || !tab.id) {
    return;
  }
  if (!tab.url.includes('chrome://')) {
    chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: parseOrders
    });
  }
});
