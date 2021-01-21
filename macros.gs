function makeStackedColumnChart(title, sheet, dataRange, chartRow, chartColumn) {
  // dataRange = spreadsheet.getRange('B2:F14')
  let chart = sheet.newChart()
  .asColumnChart()
  .addRange(dataRange)
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(1)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('useFirstColumnAsDomain', true)
  .setOption('isStacked', 'absolute')
  .setOption('title', title)
  .setXAxisTitle('')
  .setPosition(chartRow, chartColumn, 0, 0)
  .build();
  sheet.insertChart(chart);
};

