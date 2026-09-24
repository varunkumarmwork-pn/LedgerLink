// Sample review items for the Review panel until review is driven by the engine.

export const reviewItems = [
  { id: 1, tone: 'red', title: 'Other Current Assets', ref: 'D21', text: 'This year links to RMC receivable only (Note 12, row 74). Last year linked the note total, which includes GST receivable.', actions: ['Link to note total', 'Keep as is'] },
  { id: 2, tone: 'red', title: 'GODS A/C not used', ref: 'TB row 15', text: '₹3,384.00 credit under Capital Account does not reach any schedule.', actions: ['Map ledger'] },
  { id: 3, tone: 'amber', title: 'Inventories', ref: 'D16', text: 'Closing stock ₹33,98,447.26 plus Inventory Difference ₹6,23,000.00 from the TB. Confirm the difference belongs in stock.', actions: ['Approve', 'Change'] },
];
