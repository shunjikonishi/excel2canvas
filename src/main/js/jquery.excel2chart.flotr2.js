(function ($) {
	if (typeof(Flotr) === "object") {
		var defaultOptions = {
			"HtmlText" : false
		};
		$.fn.excelToChart = function(chart, op) {
			var options = $.extend({}, defaultOptions);
			if (op) {
				$.extend(options, op);
			}
			function buildChartOption() {
				var type = chart.type,
					option = chart.option,
					base = {};
				switch (type) {
					case "PIE":
						base = {
							"HtmlText" : options.HtmlText,
							"grid" : {
								"verticalLines" : false,
								"horizontalLines" : false
							},
							"xaxis" : { 
								"showLabels" : false 
							},
							"yaxis" : {
								"showLabels" : false 
							},
							"pie" : {
								"show" : true, 
								"fill" : true,
								"fillColor": null,
								"fillOpacity": 0.5,
								"explode" : 0,
								"startAngle" : 0.75
							},
							"legend" : {
								position : "se",
								backgroundColor : "#D2E8FF"
							}
						};
						break;
					case "BAR":
						var horizontal = option.bars.horizontal;
						base = {
							"HtmlText" : options.HtmlText,
							"bars" : {
								"show" : true,
								"horizontal" : false,
								"shadowSize" : 0,
								"barWidth" : 0.5
							},
							"xaxis" : {
								"min" : 0
							},
							"yaxis" : {
								"min" : 0
							},
							"grid" : {
								"verticalLines" : horizontal,
								"horizontalLines" : !horizontal
							},
							"legend" : {
								position : horizontal ? "se" : "nw",
								backgroundColor : "#D2E8FF"
							}
						};
						break;
					case "LINE":
						base = {
							"HtmlText" : options.HtmlText,
							"xaxis" : {
								"min" : 0,
								"max" : chart.labels.length + 1
							},
							"yaxis" : {
								"min" : 0
							},
							"legend" : {
								position : "nw",
								backgroundColor : "#D2E8FF"
							}
						};
						break;
					case "RADAR":
						base = {
							"HtmlText" : options.HtmlText,
							"radar" : {
								"show" : true,
								"fill" : false
							},
							"grid" : {
								"circular" : true,
								"minorHorizontalLines" : true,
								"tickColor" : "#999999"
							},
							"xaxis" : {
							},
							"yaxis" : {
								"min" : 0,
								"minorTickFreq" : 2
							},
							"legend" : {
								position : "nw",
								backgroundColor : "#D2E8FF"
							}
						};
						break;
					case "BUBBLE":
						base = {
							"HtmlText" : options.HtmlText,
							"bubbles" : {
								"show" : true
							},
							"xaxis" : {
								"min" : 0
							},
							"yaxis" : {
								"min" : 0
							},
							"legend" : {
								position : "nw",
								backgroundColor : "#D2E8FF"
							}
						};
						break;
				}
				if (chart.labels && chart.labels.length) {
					var ticks = [],
						len = chart.data.length + 1;
					for (var i=0; i<chart.labels.length; i++) {
						var n = i+1;
						if (type == "BAR" && option.bars.stacked == false) {
							n = (i * len) + (len / 2);
						}
						var tick = [n, chart.labels[i]];
						ticks.push(tick);
					}
					if (horizontal) {
						base.yaxis.ticks = ticks;
					} else {
						base.xaxis.ticks = ticks;
					}
				}
				return $.extend(true, base, option);
			}
			var chartData = chart.data,
				chartOption = buildChartOption(chart);
			Flotr.draw($(this)[0], chartData, chartOption);
		}
	}
})(jQuery);
