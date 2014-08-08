(function ($) {
	var
		BORDER_THIN                = 1,
		BORDER_MEDIUM              = 2,
		BORDER_DASHED              = 3,
		BORDER_HAIR                = 4,
		BORDER_THICK               = 5,
		BORDER_DOUBLE              = 6,
		BORDER_DOTTED              = 7,
		BORDER_MEDIUM_DASHED       = 8,
		BORDER_DASH_DOT            = 9,
		BORDER_MEDIUM_DASH_DOT     = 10,
		BORDER_DASH_DOT_DOT        = 11,
		BORDER_MEDIUM_DASH_DOT_DOT = 12,
		BORDER_SLANTED_DASH_DOT    = 13,
		context;
	
	function hasTooltip() {
		return !!$.fn.tooltip
	}
	function fillStyle(data, fill) {
		var back = fill.back,
			fore = fill.fore,
			pattern = fill.pattern;
		if (fill.styleRef) {
			var styles = data.styles[fill.styleRef].split("|");
			back = styles[0];
			fore = styles[1];
			pattern = styles[2];
		}
		//ToDo back, pattern
		if (fore) {
			context.fillStyle = fore;
			context.fillRect(fill.p[0], fill.p[1], fill.p[2], fill.p[3]);
		}
	}
	function drawLine(line) {
		var kind = line.kind ? line.kind : 1,
			w = 1,
			x1 = line.p[0],
			y1 = line.p[1],
			x2 = line.p[2],
			y2 = line.p[3],
			horizontal = y1 == y2;
		
		if (kind == BORDER_MEDIUM || 
		    kind == BORDER_MEDIUM_DASHED || 
		    kind == BORDER_MEDIUM_DASH_DOT ||
		    kind == BORDER_MEDIUM_DASH_DOT_DOT) 
		{
			w = 2;
		} else if (kind == BORDER_THICK) {
			w = 3;
		}
		context.lineWidth = w;
		if (w == 1 || w == 3) {
			if (horizontal) {
				y1 = y1 == 0 ? 0.5 : y1 - 0.5;
				y2 = y2 == 0 ? 0.5 : y2 - 0.5;
			} else {
				x1 = x1 == 0 ? 0.5 : x1 - 0.5;
				x2 = x2 == 0 ? 0.5 : x2 - 0.5;
			}
		} else {//w == 2
			if (horizontal && y1 == 0) {
				y1 = 1;
				y2 = 1;
			} else if (!horizontal && x1 == 0) {
				x1 = 1;
				x2 = 1;
			}
		}
		if (line.color) {
			context.strokeStyle = line.color;
		} else {
			context.strokeStyle = "#000000";
		}
		if (kind == BORDER_DOUBLE) {
			context.beginPath();
			if (horizontal) {
				context.moveTo(x1, y1 - 1);
				context.lineTo(x2, y2 - 1);
				context.moveTo(x1, y1 + 1);
				context.lineTo(x2, y2 + 1);
			} else {
				context.moveTo(x1 - 1, y1);
				context.lineTo(x2 - 1, y2);
				context.moveTo(x1 + 1, y1);
				context.lineTo(x2 + 1, y2);
			}
			context.stroke();
			context.closePath();
		} else if (kind == BORDER_DASHED ||
		           kind == BORDER_HAIR ||
		           kind == BORDER_DOTTED ||
		           kind == BORDER_MEDIUM_DASHED)
		{
			var bw, sw;
			switch (kind) {
				case BORDER_DASHED:
					bw = 4; sw = 2;
					break;
				case BORDER_HAIR:
					bw = 2; sw = 2;
					break;
				case BORDER_DOTTED:
					bw = 1; sw = 1;
					break;
				case BORDER_MEDIUM_DASHED:
					bw = 8; sw = 3;
					break;
			}
			
			var bar = true;
			context.beginPath();
			context.moveTo(x1, y1);
			if (horizontal) {
				var y = y1,
					cx = x1,
					ex = x2;
				while (cx < ex) {
					var nx = bar ? bw : sw;
					cx += nx;
					if (cx > ex) {
						cx = ex;
					}
					if (bar) {
						context.lineTo(cx, y);
					} else {
						context.moveTo(cx, y);
					}
					bar = !bar;
				}
			} else {
				var x = x1,
					cy = y1,
					ey = y2;
				while (cy < ey) {
					var ny = bar ? bw : sw;
					cy += ny;
					if (cy > ey) {
						cy = ey;
					}
					if (bar) {
						context.lineTo(x, cy);
					} else {
						context.moveTo(x, cy);
					}
					bar = !bar;
				}
			}
			context.stroke();
			context.closePath();
		} else {
			context.beginPath();
			context.moveTo(x1, y1);
			context.lineTo(x2, y2);
			context.stroke();
			context.closePath();
		}
	}
	$.fn.excelToCanvas = function(data, mergeData) {
		var holder, canvas;
		if (this[0].tagName == "canvas") {
			canvas = this;
			holder = canvas.parent();
		} else {
			holder = this;
			canvas = holder.find("canvas");
			if (canvas.length == 0) {
				canvas = $("<canvas style='position:absolute;left:0;top:0;z-index:0'></canavas>").appendTo(holder);
			}
		}
		if (typeof FlashCanvas !== "undefined") {
			FlashCanvas.initElement(canvas[0]);
		}
		
		canvas.attr("width", data.width).attr("height", data.height);
		holder.css({
			"width" : data.width,
			"height" : data.height
		});
		context = canvas[0].getContext("2d");
		context.fillStyle = "white";
		context.fillRect(0, 0, data.width, data.height);
		if (data.font) {
			context.font = data.font;
		}
		
		if (data.fills) {
			for (var i=0; i<data.fills.length; i++) {
				var fill = data.fills[i];
				fillStyle(data, fill);
			}
		}
		if (data.lines) {
			for (var i=0; i<data.lines.length; i++) {
				var line = data.lines[i];
				drawLine(line);
			}
		}
		
		if (data.strs) {
			for (var i=0; i<data.strs.length; i++) {
				var str = data.strs[i],
					style = str.style ? str.style : data.styles[str.styleRef],
					span = $("<span style='" + style + "'></span>");
				if (str.link) {
					var link = $("<a target='_blank'></a>");
					link.append(str.text);
					link.attr("href", str.link);
					span.append(link);
				} else {
					span.append(str.text);
				}
				var align = str.align;
				if (align) {
					span.addClass("cell-a" + align[0]);
					span.addClass("cell-v" + align[1]);
					if (align[1] == "c" || align[1] == "j") {
						var n = str.text.split("<br>").length;
						if (n > 1) {
							span.css("margin-top", "-" + (n / 2) + "em");
						}
					}
				}
				if (str.rawdata) {
					span.attr("data-raw", str.rawdata);
				}
				var div = $("<div id='" + str.id + "' class='cell'></div>").append(span);
				div.css({
					"left" : str.p[0],
					"top" : str.p[1],
					"width" : str.p[2],
					"height" : str.p[3]
				});
				if (str.clazz) {
					div.addClass(str.clazz);
				}
				if (str.dataAttrs) {
					$.each(str.dataAttrs, function(key, value) {
						div.attr("data-" + key, value);
					});
				}
				if (str.comment && hasTooltip()) {
					context.strokeStyle = "red";
					context.fillStyle = "red";
					
					context.beginPath();
					context.moveTo(str.p[0] + str.commentWidth - 4,  str.p[1]);
					context.lineTo(str.p[0] + str.commentWidth, str.p[1]);
					context.lineTo(str.p[0] + str.commentWidth, str.p[1] + 4);
					context.lineTo(str.p[0] + str.commentWidth - 4, str.p[1]);
					context.stroke();
					context.fill();
					context.closePath();
					div.tooltip({
						"title" : str.comment,
						"html" : true
					});
				}
				holder.append(div);
			}
		}
		
		if (data.pictures) {
			for (var i=0; i<data.pictures.length; i++) {
				var pic = data.pictures[i],
					img = $("<img class='excel-img'/>");
				img.attr("src", pic.uri);
				img.css({
					"left" : pic.p[0],
					"top" : pic.p[1],
					"width" : pic.p[2],
					"height" : pic.p[3]
				});
				holder.append(img);
			}
		}
		if (data.charts && $.fn.excelToChart) {
			for (var i=0; i<data.charts.length; i++) {
				var chart = data.charts[i],
					options = {},
					chartDiv = $("<div class='excel-chart'></div>");
				chartDiv.css({
					"left" : chart.p[0],
					"top" : chart.p[1],
					"width" : chart.p[2],
					"height" : chart.p[3]
				});
				holder.append(chartDiv);
				if (data.font) {
					options.HtmlText = true;
				}
				chartDiv.excelToChart(chart.chart, options);
				chartDiv.css("position", "absolute");
			}
		}
		if (mergeData) {
			for (var key in mergeData) {
				var div = $("#" + key);
				if (div.length > 0) {
					var userData = mergeData[key];
					if (typeof (userData) === "string") {
						div.find("span").text(userData);
					} else if (userData.element) {
						var el = $(document.createElement(userData.element));
						for (attr in userData) {
							if (attr != "element") {
								el.attr(attr, userData[attr]);
							}
						}
						div.css("overflow", "visible").empty().append(el);
					}
				}
			}
		}
		return this;
	}
})(jQuery);
