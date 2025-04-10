/**
  SmartPost - Modified version of an Autodesk Fusion 360 post processor
  Original postprocessor (C) 2012–2021 by Autodesk, Inc.
  Modifications (C) 2025 by MaestroFusion360

  This modified post is part of the SmartPost project:
  https://github.com/MaestroFusion360/

  Licensed under the Apache License, Version 2.0

  Original Autodesk Revision: 43793 d40578e2e144c6eec8d617d28994bee231ce87cd
  Original Date: 2022-05-04
*/

description = "SmartPost: XML Cutter Location Data";
vendor = "SmartPost by MaestroFusion360";
vendorUrl = "https://github.com/MaestroFusion360/";
legal = "Modifications (C) 2025 by MaestroFusion360. Based on original work by Autodesk, Inc.";
certificationLevel = 0;

longDescription = "Example post illustrating how to convert the toolpath into XML.";

capabilities = CAPABILITY_INTERMEDIATE;
extension = "xml";
setCodePage("utf-8");
var expanding = false;
allowHelicalMoves = true;
allowedCircularPlanes = undefined; // allow any circular motion

properties = {
  useTimeStamp: {
    title      : "Time stamp",
    description: "Specifies whether to output a time stamp.",
    group      : "preferences",
    type       : "boolean",
    value      : false,
    scope      : "post"
  },
  highAccuracy: {
    title      : "High accuracy",
    description: "Specifies short (no) or long (yes) numeric format.",
    group      : "preferences",
    type       : "boolean",
    value      : true,
    scope      : "post"
  }
};

var mainFormat = createFormat({decimals:6, forceDecimal:false});
var ijkFormat = createFormat({decimals:9, forceDecimal:false});

var feedOutput = createVariable({format:mainFormat});

var mapRCTable = new Table(
  [" compensation='off'", " compensation='left'", "", " compensation='right'"],
  {initial:RADIUS_COMPENSATION_OFF},
  "Invalid radius compensation"
);

function toPos(x, y, z) {
  return mainFormat.format(x) + " " + mainFormat.format(y) + " " + mainFormat.format(z);
}

function toVec(x, y, z) {
  return ijkFormat.format(x) + " " + ijkFormat.format(y) + " " + ijkFormat.format(z);
}

function toFeed(feed) {
  var f = feedOutput.format(feed);
  return f ? (" feed='" + f + "'") : "";
}

function toRC(radiusCompensation) {
  // return mapRCTable.lookup(radiusCompensation);
  switch (radiusCompensation) {
  case RADIUS_COMPENSATION_OFF:
    return " compensation='off'";
  case RADIUS_COMPENSATION_LEFT:
    return " compensation='left'";
  case RADIUS_COMPENSATION_RIGHT:
    return " compensation='right'";
  }
  return "";

}

function escapeChar(c) {
  switch (c) {
  case "<":
    return "&lt;";
  case ">":
    return "&gt;";
  case "&":
    return "&amp;";
  case "'":
    return "&apos;";
  case "\"":
    return "&quot;";
  }
  return c; // should never happen
}

function escapeXML(unescaped) {
  return unescaped.replace(/[<>&'"]/g, escapeChar);
}

function makeValue(value) {
  if (typeof value == "string") {
    return escapeXML(value);
  } else if (typeof value == "number" && !isNaN(value)) {
    return mainFormat.format(value);
  } else if (value instanceof Array || (typeof value == "object" && value.x !== undefined)) {
    var parts = [];
    if (value instanceof Array) {
      for (var i = 0; i < value.length; ++i) {
        parts.push(makeValue(value[i]));
      }
    } else {
      parts.push(makeValue(value.x));
      parts.push(makeValue(value.y));
      parts.push(makeValue(value.z));
    }
    return parts.join(", ");
  } else {
    error("makeValue: invalid value, expected number, string, or vector/array, got: " + value);
    return "";
  }
}

function onOpen() {
  writeln("<?xml version='1.0' encoding='utf-8' standalone='yes'?>");
  writeln("<nc xmlns='http://www.hsmworks.com/xml/2008/nc' version='1.0'>");
  writeln("<!-- http://cam.autodesk.com -->");
  if (getProperty("useTimeStamp")) {
    var d = new Date();
    writeln("<meta><date timestamp='" + (d.getTime() * 1000) + "'/></meta>");
  }

  if (!getProperty("highAccuracy")) {
    mainFormat = createFormat({decimals:4, forceDecimal:true});
    ijkFormat = createFormat({decimals:7, forceDecimal:true});
    feedOutput = createVariable({format:mainFormat});
  }
}

function onComment(text) {
  writeln("<comment>" + escapeXML(text) + "</comment>");
}

function attr(name, value) {
  return name + "='" + value + "'";
}

function onSection() {
  var u = (unit == IN) ? "inches" : "millimeters";
  var o = toPos(currentSection.workOrigin.x, currentSection.workOrigin.y, currentSection.workOrigin.z);
  var p = [];
  for (var i = 0; i < 9; ++i) {
    p.push(currentSection.workPlane.getElement(i / 3, i % 3));
  }

  writeln("<context " + attr("unit", u) + " " + attr("origin", o) + " " + attr("plane", p.join(" ")) + " " + attr("work-offset", currentSection.workOffset) + "/>");

  if (currentSection.isPatterned && currentSection.isPatterned()) {
    var patternId = currentSection.getPatternId();
    var sections = [];
    for (var i = 0; i < getNumberOfSections(); ++i) {
      var section = getSection(i);
      if (section.getPatternId() == patternId) {
        sections.push(section.getId());
      }
    }
    writeln("<!-- Pattern ID: " + patternId + ", instances: " + sections.join(", ") + " -->");
  }

  var type = getToolTypeName(tool.type);
  var n = mainFormat.format(tool.number);
  var d = mainFormat.format(tool.diameter);
  var cr = mainFormat.format(tool.cornerRadius);
  var ta = mainFormat.format(tool.taperAngle);
  var fl = mainFormat.format(tool.fluteLength);
  var sl = mainFormat.format(tool.shoulderLength);
  var sd = mainFormat.format(tool.shaftDiameter);
  var bl = mainFormat.format(tool.bodyLength);
  var tp = mainFormat.format(tool.threadPitch);
  var _do = mainFormat.format(tool.diameterOffset);
  var lo = mainFormat.format(tool.lengthOffset);
  var sr = mainFormat.format(tool.spindleRPM);

  var COOLANT_NAMES = ["disabled", "flood", "mist", "tool", "air", "air through tool"];
  var coolant = COOLANT_NAMES[tool.coolant];

  writeln("<tool type='" + type + "' number='" + n + "' diameter='" + d + "' corner-radius='" + cr + "' taper-angle='" + ta + "' flute-length='" + fl + "' shoulder-length='" + sl + "' body-length='" + bl + "' shaft-diameter='" + sd + "' thread-pitch='" + tp + "' diameter-offset='" + _do + "' length-offset='" + lo + "' spindle-rpm='" + sr + "' coolant='" + coolant + "'>");
  // writeln("<!-- DEBUG: Tool Type = " + type + " -->");
  var holder = tool.holder;
  if (holder) {
    writeln("<holder>");
    for (var i = 0; i < holder.getNumberOfSections(); ++i) {
      var d = mainFormat.format(holder.getDiameter(i));
      var l = mainFormat.format(holder.getLength(i));
      writeln("<section diameter='" + d + "' length='" + l + "'/>");
    }
    writeln("</holder>");
  }
  writeln("</tool>");

  writeln("<section>");

  feedOutput.reset();
}

function onParameter(name, value) {
  var type = "float";
  if (typeof value  == "string") {
    type = "string";
  } else if ((value % 1) == 0) {
    type = "integer";
  }
  writeln("<parameter name='" + escapeXML(name) + "' value='" + makeValue(value) + "' type='" + type + "'/>");
}

function onDwell(seconds) {
  writeln("<dwell seconds='" + mainFormat.format(seconds) + "'/>");
}

function onCyclePoint(x, y, z) {
  expanding = true;
  expandCyclePoint(x, y, z);
  expanding = false;
}

function onRapid(x, y, z) {
  writeln("<rapid to='" + toPos(x, y, z) + "'" + toRC(radiusCompensation) + "/>");
  feedOutput.reset();
}

function onLinear(x, y, z, feed) {
  writeln("<linear to='" + toPos(x, y, z) + "'" + toFeed(feed) + toRC(radiusCompensation) + "/>");
}

function onRapid5D(x, y, z, dx, dy, dz) {
  writeln("<rapid5d to='" + toPos(x, y, z) + "' axis='" + toPos(dx, dy, dz) + "'/>");
  previousFeed = undefined;
}

function onLinear5D(x, y, z, dx, dy, dz, feed) {
  writeln("<linear5d to='" + toPos(x, y, z) + "' axis='" + toVec(dx, dy, dz) + "'" + toFeed(feed) + "/>");
}

function onCircular(clockwise, cx, cy, cz, x, y, z, feed) {
  var n = getCircularNormal();
  var block = "";
  var big = getCircularSweep() > Math.PI;
  if (big) {
    block += "circular";
  } else {
    block += isClockwise() ? "arc-cw" : "arc-ccw";
  }
  block += " to='" + toPos(x, y, z) + "'";
  block += " center='" + toPos(cx, cy, cz) + "'";
  if ((n.x != 0) || (n.y != 0) || (n.z != 1)) {
    block += " normal='" + toVec(n.x, n.y, n.z) + "'";
  }
  if (big) {
    block += " sweep='" + mainFormat.format(getCircularSweep()) + "'";
  }
  block += toFeed(feed);
  block += toRC(radiusCompensation);
  writeln("<" + block + "/>");
}

function onCommand() {
  writeln("<command/>");
}

function onSectionEnd() {
  writeln("</section>");
}

function onClose() {
  writeln("</nc>");
}

function setProperty(property, value) {
  properties[property].current = value;
}


