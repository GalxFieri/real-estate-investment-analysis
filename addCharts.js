/**
 * Add charts to the dashboard sheet
 * This function should be called after setupDashboard()
 */
function addCharts(dashboardSheet, mainSheet) {
  // 1. Monthly Cash Flow Breakdown Pie Chart
  var cashFlowChart = dashboardSheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(mainSheet.getRange("A24:A28"))   // Labels: expense categories
    .addRange(mainSheet.getRange("B24:B28"))   // Values: expense amounts
    .setPosition(15, 1, 0, 0)
    .setOption('title', 'Monthly Expenses Breakdown')
    .setOption('pieSliceText', 'percentage')
    .setOption('legend', {position: 'right'})
    .setOption('width', 500)
    .setOption('height', 300)
    .build();
  
  dashboardSheet.insertChart(cashFlowChart);
  
  // 2. Appreciation Rate Sensitivity Chart (Line chart)
  var appreciationChart = dashboardSheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(mainSheet.getRange("A56:A59"))   // Labels: appreciation rates
    .addRange(mainSheet.getRange("B56:C59"))   // Values: 5-yr and 10-yr property values
    .setPosition(15, 5, 0, 0)
    .setOption('title', 'Property Value Projection by Appreciation Rate')
    .setOption('hAxis', {title: 'Appreciation Rate'})
    .setOption('vAxis', {title: 'Projected Value ($)'})
    .setOption('series', {
      0: {targetAxisIndex: 0, labelInLegend: '5-Year Value'},
      1: {targetAxisIndex: 0, labelInLegend: '10-Year Value'}
    })
    .setOption('width', 500)
    .setOption('height', 300)
    .build();
  
  dashboardSheet.insertChart(appreciationChart);
  
  // 3. Interest Rate Sensitivity Chart (Column chart)
  var interestRateChart = dashboardSheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(mainSheet.getRange("A48:A52"))   // Labels: interest rates
    .addRange(mainSheet.getRange("C48:C52"))   // Values: monthly cash flow
    .setPosition(30, 1, 0, 0)
    .setOption('title', 'Monthly Cash Flow by Interest Rate')
    .setOption('hAxis', {title: 'Interest Rate'})
    .setOption('vAxis', {title: 'Monthly Cash Flow ($)'})
    .setOption('width', 500)
    .setOption('height', 300)
    .setOption('colors', ['#4285F4'])
    .build();
  
  dashboardSheet.insertChart(interestRateChart);
  
  // 4. Scenario Comparison Chart (Bar chart)
  var scenarioChart = dashboardSheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(mainSheet.getRange("A65:A68"))   // Labels: metrics
    .addRange(mainSheet.getRange("B65:D68"))   // Values: scenarios
    .setPosition(30, 5, 0, 0)
    .setOption('title', 'Scenario Comparison')
    .setOption('vAxis', {title: 'Metric'})
    .setOption('hAxis', {title: 'Value'})
    .setOption('width', 500)
    .setOption('height', 300)
    .setOption('colors', ['#4285F4', '#34A853', '#FBBC05'])
    .setOption('legend', {position: 'top'})
    .build();
  
  dashboardSheet.insertChart(scenarioChart);

  // 5. ROI Gauge Chart
  var roiGaugeChart = dashboardSheet.newChart()
    .setChartType(Charts.ChartType.GAUGE)
    .addRange(mainSheet.getRange("A43"))      // Label: ROI %
    .addRange(mainSheet.getRange("B43"))      // Value: Current ROI
    .setPosition(45, 3, 0, 0)
    .setOption('title', 'Return on Investment')
    .setOption('min', 0)
    .setOption('max', 30)
    .setOption('greenFrom', 15)
    .setOption('greenTo', 30)
    .setOption('yellowFrom', 5)
    .setOption('yellowTo', 15)
    .setOption('redFrom', 0)
    .setOption('redTo', 5)
    .setOption('width', 400)
    .setOption('height', 200)
    .build();
  
  dashboardSheet.insertChart(roiGaugeChart);
}