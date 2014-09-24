Excel2Canvas
-----------------
-----------------

Overview
--------
Excel2Canvas is a library to display an Excel file on the web browser.

Its source codes consists of two parts.

1. **Java.** It is used to convert Excel file to JSON.
2. **JavaScript.** It is used to draw Excel view on the web browser.

JavaDoc
-------
http://oss.flect.co.jp/apidocs/excel2canvas/index.html

Usage
-----
At first, read an excel file and convert it to JSON string.

    ExcelToCanvasBuilder builder = new ExcelToCanvasBuilder();
    builder.setIncludeComment(true);//If need display comments.
    builder.setIncludeChart(true);//If need display charts.(Flotr2 is required.)
    builder.setIncludePicture(true);//If need display picture.
    String json = builder.build(new File("Book1.xlsx"), "Sheet1").toJson();
    
Next, embed JSON string to HTML, and apply jQuery plugin method to a div element that holding a canvas element.

    <script src="//ajax.googleapis.com/ajax/libs/jquery/1.7.1/jquery.min.js" type="text/javascript"></script>
    
    <link rel="stylesheet" type="text/css" media="screen,print" href="jquery.excel2canvas.css" />
    <script type="text/javascript" language="javascript"  src="flotr2.js"></script>
    <script type="text/javascript" language="javascript"  src="jquery.excel2canvas.min.js"></script>
    <script type="text/javascript" language="javascript"  src="jquery.excel2chart.flotr2.min.js"></script>

    <script>
    $(function() {
    	var data = ${excel.raw()};//Get the JSON string.(Template engine syntax.)
    	$("#canvasHolder").excelToCanvas(data).css("width", $("#canvas").width());
    });
    </script>
    
    ...
    
    <div id="canvasHolder">
    	<canvas id="canvas"></canvas>
    </div>

Supported Excel features
------------------------
* Background color
    * Simple color is supported.
    * Patterns and graduations are not.
* Lines
    * Simple line, Double line, and Dash line are supported.
    * Not supported lines are drawn as simple line.
* Pictures
    * png and jpeg are supported.
    * All of decorations are not supported.(e.g. border, rotation)
* String
    * Font : Depends on browser and OS font.
    * Format ： Except localized format.(Starts with "*")
    * Horizontal alignment： Left, Right, Center
    * Vertica alignmentl：Top, Bottom, Middle
    * HyperLink : Support.
    * Merged cell : Support.
    * Rich string: Not support.
* Comment
    * Support.
* Formula
    * Depends on [Apache POI](http://poi.apache.org/)
* Chart
    * Partial support.
    * xlsx only
    * Support bar chart, pie chart, line chart, and radar chart.
    * Properties are not considered.(e.g. Colors, marks of line chart, position of legend)
    
Dependencies
------------
The java library is highly dependent on [Apache POI](http://poi.apache.org/) and other apache libraries.  
And includes [Google gson](https://code.google.com/p/google-gson/) to parse JSON string.


The javascript library is implemented as a [jQuery](http://jquery.com/) plugin.  
If you want to display chart, you must include [Flotr2](http://humblesoftware.com/flotr2/).  
If you want to display comment, you must include [Bootstrap](http://twitter.github.com/bootstrap/).  

Samples
-------
[ExcelNote](https://excelnote.herokuapp.com/) is an application that uses this library.  
If the Excel file is attached in Evernote, it allows you to display and edit it on the web browser.

And if that note is shared with the world, it is published to everyone.

Followings are some of its samples.(Japanese only)

- [Time schedule](http://excelnote.herokuapp.com/share/note/s91/90ae165a-18b7-4879-a667-6ad15bbcd57b/5e1a3c243456d0e3daf8bd42005a22e0?theme=humanity)
- [The result of a baseball game](http://excelnote.herokuapp.com/share/note/s91/e94bd16f-465a-4a24-b71e-dff906cf3395/67079ce23db9c8af6df06b33d12c8e70?theme=sunny)
- [Purchase order](https://excelnote.herokuapp.com/share/excel/s91/09880d80-43bd-4728-9f8f-300b84a3a32c/151561d3185ea4c1dc7fa3ab3f2db653?sheet=%E7%99%BA%E6%B3%A8%E6%9B%B8&theme=redmond)  
  It is able to print as is from web browser.

Vesion
------
Current version is 1.2.3

Install
-------
You can install this library from FLECT maven repository.

sbt  

```
resolvers += "FLECT" at "http://flect.github.io/maven-repo/"

libraryDependencies += "jp.co.flect" % "excel2canvas" % "1.2.3"
```

bower  

```
bower install https://github.com/shunjikonishi/excel2canvas.git
```

License
-------
MIT

