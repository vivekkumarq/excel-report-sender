// What gets sent, in this exact order. Names are the workbook's named ranges (Excel's Name Box list).
//   as: 'text'  -> sent as a WhatsApp message (the cells' text, one line per row)
//   as: 'image' -> PNG picture of the range
//   as: 'pdf'   -> PDF; give several names to get one PDF with a page per range
// Consecutive items of the same kind are shared together, so this list becomes 4 WhatsApp sends.
const images = ['Alok_kr_Singh_Sir', 'Alok_Kr_Singh_Sir_01', '_ccl_coal', '_ccl_obr', '_ccl_offtake', '_ccl_rake',
  'AREAWISE_STOCK', 'CMD_Report_01', 'CMD_Report_02', 'CMD_Report_06', '_area_wise_mine_wise', '_dep_obr_ccl',
  '_Projection_ccl', 'CCL_COMPARATIVE_REPORT', 'Rainfall_Report', 'CCL_GMs', 'GRAPHICAL_Perf_Area',
  'Yearly_Cap_Utlzn_Report', 'Report_BookPlusSiding_Stock', 'Top_11', 'Top_14', 'Top_20', 'Washery_Report'];
const pdfs = ['A_C_A_Report', 'CCL_Summary', 'DateWise_Areawise_Report', 'Departmental_Only_Project',
  'Last_6_Yr_Mthly_Report', 'Quick_Report_110_CMD', 'UG_Qck_Rpt_Coal'];

export const REPORTS = [
  { name: 'SMS_To_IMC', as: 'text' },
  ...images.map(name => ({ name, as: 'image' })),
  ...pdfs.map(name => ({ name, as: 'pdf' })),
  { name: ['_ccl_coal', '_ccl_obr', '_ccl_offtake', '_ccl_rake'], as: 'pdf', file: 'CCL_Coal_OBR_Offtake_Rake' },
];
