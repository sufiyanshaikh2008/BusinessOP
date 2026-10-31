const APP_NAME = 'BusinessOP';
const PROP_DB  = 'BOP_DB_ID';

const SCHEMA = {
  Products:     ['id','name','price','cost','stock','min','updated'],
  Customers:    ['id','name','phone','updated'],
  Vendors:      ['id','name','phone','updated'],
  Invoices:     ['id','no','date','cust','total','paid','updated'],
  InvoiceItems: ['invoiceId','productId','name','qty','price','cost'],
  Purchases:    ['id','vendor','productId','name','qty','cost','date'],
  Payments:     ['id','invoiceId','amount','date','mode'],
  Meta:         ['key','value']
};
