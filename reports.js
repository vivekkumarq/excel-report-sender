// What gets sent, in this exact order. Names are the workbook's named ranges (Excel's Name Box list).
//   as: 'text'  -> sent as a WhatsApp message (the cells' text, one line per row)
//   as: 'image' -> PNG picture of the range
//   as: 'pdf'   -> PDF named `file`, laid out with `page` (see print.js). Several names make one PDF with the
//                  pictures stacked `perPage` to a page.
// Consecutive items of the same kind are sent together, so this list becomes 3 sends: text, images, PDFs.
const images = ['Alok_kr_Singh_Sir', 'Alok_Kr_Singh_Sir_01', '_ccl_coal', '_ccl_obr', '_ccl_offtake', '_ccl_rake',
  'AREAWISE_STOCK', 'CMD_Report_01', 'CMD_Report_02', 'CMD_Report_06', '_area_wise_mine_wise', '_dep_obr_ccl',
  '_Projection_ccl', 'CCL_COMPARATIVE_REPORT', 'Rainfall_Report', 'CCL_GMs', 'GRAPHICAL_Perf_Area',
  'Yearly_Cap_Utlzn_Report', 'Report_BookPlusSiding_Stock', 'Top_11', 'Top_14', 'Top_20', 'Washery_Report'];

// File names and page setup copied from the sample PDFs (Excel, Microsoft Print to PDF): US Letter, margins in
// points as measured on each sample, the range fitted to the page width ('page': to the whole page). Rows run
// on to further pages where a page is full, and at the sheet's own page breaks.
const m = (l, t, r, b) => ({ l, t, r, b });
const pdfs = [
  ['A_C_A_Report', 'ACA Report', 'portrait', m(51, 25, 51, 25)],
  ['CCL_Summary', 'CCL Summary', 'portrait', m(13.5, 15, 15, 15)],
  ['DateWise_Areawise_Report', 'Datewise Areawise', 'landscape', m(14.7, 16.8, 25.7, 25)],
  ['Departmental_Only_Project', 'Departmental', 'portrait', m(50, 26, 83, 26), 'page'],
  ['Last_6_Yr_Mthly_Report', 'Last 6 Years', 'landscape', m(14, 14, 15.9, 16)],
  ['Quick_Report_110_CMD', 'Quick Report', 'portrait', m(2, 0, 10, 10), 'page'],
  ['UG_Qck_Rpt_Coal', 'UG Report', 'landscape', m(1.3, 0.5, 5.3, 13)],
];

export const REPORTS = [
  { name: 'SMS_To_IMC', as: 'text' },
  ...images.map(name => ({ name, as: 'image' })),
  ...pdfs.map(([name, file, orientation, margins, fit = 'width']) => ({ name, as: 'pdf', file, page: { paper: 'letter', orientation, margins, fit } })),
  // like the sample made in Word: A4, 1-inch margins, two reports per page at full text width
  { name: ['_ccl_coal', '_ccl_obr', '_ccl_offtake', '_ccl_rake'], as: 'pdf', file: 'CCL Performance Report', page: { paper: 'a4', orientation: 'portrait', margins: 72, perPage: 2 } },
];
