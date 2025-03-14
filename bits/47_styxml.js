/* 18.8.21 fills CT_Fills */
function parse_fills(t, opts) {
	styles.Fills = [];
	var fill = {};
	t[0].match(tagregex).forEach(function(x) {
		var y = parsexmltag(x);
		switch(y[0]) {
			case '<fills': case '<fills>': case '</fills>': break;

			/* 18.8.20 fill CT_Fill */
			case '<fill>': break;
			case '</fill>': styles.Fills.push(fill); fill = {}; break;

			/* 18.8.32 patternFill CT_PatternFill */
			case '<patternFill':
				if(y.patternType) fill.patternType = y.patternType;
				break;
			case '<patternFill/>': case '</patternFill>': break;

			/* 18.8.3 bgColor CT_Color */
			case '<bgColor':
				if(!fill.bgColor) fill.bgColor = {};
				if(y.indexed) fill.bgColor.indexed = parseInt(y.indexed, 10);
				if(y.theme) fill.bgColor.theme = parseInt(y.theme, 10);
				if(y.tint) fill.bgColor.tint = parseFloat(y.tint);
				/* Excel uses ARGB strings */
				if(y.rgb) fill.bgColor.rgb = y.rgb.substring(y.rgb.length - 6);
				break;
			case '<bgColor/>': case '</bgColor>': break;

			/* 18.8.19 fgColor CT_Color */
			case '<fgColor':
				if(!fill.fgColor) fill.fgColor = {};
				if(y.theme) fill.fgColor.theme = parseInt(y.theme, 10);
				if(y.tint) fill.fgColor.tint = parseFloat(y.tint);
				/* Excel uses ARGB strings */
				if(y.rgb) fill.fgColor.rgb = y.rgb.substring(y.rgb.length - 6);
				break;
			case '<fgColor/>': case '</fgColor>': break;

			default: if(opts.WTF) throw 'unrecognized ' + y[0] + ' in fills';
		}
	});
}

function write_fills(fills) {
	var o = [
		'<fills count="' + (fills ? 1 + fills.length : 1) + '">',
		'<fill>' + writextag('patternFill', null, {patternType: "none"}) + '</fill>'
	];
	if (fills) {
		for (var i = 0; i < fills.length; i++) {
			var fill = JSON.parse(fills[i]);
			o.push('<fill><patternFill patternType="'+fill.patternType+'">' +
				writextag("fgColor", null, fill.fgColor) +
				writextag("bgColor", null, fill.bgColor) +
				'</patternFill></fill>');
		}
	}
	o[o.length] = ("</fills>");
	return o.join("");
}

function write_fonts(fonts) {
	var o = [
		'<fonts count="' + (fonts ? 1 + fonts.length : 1) + '">',
		'<font><sz val="12"/><color theme="1"/><name val="Calibri"/><family val="2"/><charset val="1"/></font>'
	];
	if (fonts) {
		for (var i = 0; i < fonts.length; i++) {
			var font = JSON.parse(fonts[i]);
			o.push('<font><sz val="12"/><name val="Calibri"/><family val="2"/><charset val="1"/>' +
				(font.color ? writextag("color", null, {rgb: font.color}) : "") +
				(font.bold ? writextag("b", null, {val: true}) : "") +
				'</font>');
		}
	}
	o[o.length] = ("</fonts>");
	return o.join("");
}

function write_borders(borders) {
	var o = [
		'<borders count="' + (borders ? 1 + borders.length : 1) + '">',
		'<border diagonalUp="false" diagonalDown="false"><left/><right/><top/><bottom/><diagonal/></border>'
	];
	if (borders) {
		for (var i = 0; i < borders.length; i++) {
			var border = JSON.parse(borders[i]);
			o.push('<border diagonalUp="' + border.diagonalUp + '" diagonalDown="' + border.diagonalDown + '">' +
				writextag("left", border.left ? writextag("color", null, border.left.color) : null, border.left ? {style: border.left.style} : null) +
				writextag("right", border.right ? writextag("color", null, border.right.color) : null, border.right ? {style: border.right.style} : null) +
				writextag("top", border.top ? writextag("color", null, border.top.color) : null, border.top ? {style: border.top.style} : null) +
				writextag("bottom", border.bottom ? writextag("color", null, border.bottom.color) : null, border.bottom ? {style: border.bottom.style} : null) +
				writextag("diagonal", border.diagonal ? writextag("color", null, border.diagonal.color) : null, border.diagonal ? {style: border.diagonal.style} : null) +
				'</border>');
		}
	}
	o[o.length] = ("</borders>");
	return o.join("");
}

/* 18.8.31 numFmts CT_NumFmts */
function parse_numFmts(t, opts) {
	styles.NumberFmt = [];
	var k = keys(SSF._table);
	for(var i=0; i < k.length; ++i) styles.NumberFmt[k[i]] = SSF._table[k[i]];
	var m = t[0].match(tagregex);
	for(i=0; i < m.length; ++i) {
		var y = parsexmltag(m[i]);
		switch(y[0]) {
			case '<numFmts': case '</numFmts>': case '<numFmts/>': case '<numFmts>': break;
			case '<numFmt': {
				var f=unescapexml(utf8read(y.formatCode)), j=parseInt(y.numFmtId,10);
				styles.NumberFmt[j] = f; if(j>0) SSF.load(f,j);
			} break;
			default: if(opts.WTF) throw 'unrecognized ' + y[0] + ' in numFmts';
		}
	}
}

function write_numFmts(NF, opts) {
	var o = ["<numFmts>"];
	[[5,8],[23,26],[41,44],[63,66],[164,392]].forEach(function(r) {
		for(var i = r[0]; i <= r[1]; ++i) if(NF[i] !== undefined) o[o.length] = (writextag('numFmt',null,{numFmtId:i,formatCode:escapexml(NF[i])}));
	});
	if(o.length === 1) return "";
	o[o.length] = ("</numFmts>");
	o[0] = writextag('numFmts', null, { count:o.length-2 }).replace("/>", ">");
	return o.join("");
}

/* 18.8.10 cellXfs CT_CellXfs */
function parse_cellXfs(t, opts) {
	styles.CellXf = [];
	t[0].match(tagregex).forEach(function(x) {
		var y = parsexmltag(x);
		switch(y[0]) {
			case '<cellXfs': case '<cellXfs>': case '<cellXfs/>': case '</cellXfs>': break;

			/* 18.8.45 xf CT_Xf */
			case '<xf': delete y[0];
				if(y.numFmtId) y.numFmtId = parseInt(y.numFmtId, 10);
				if(y.fillId) y.fillId = parseInt(y.fillId, 10);
				styles.CellXf.push(y); break;
			case '</xf>': break;

			/* 18.8.1 alignment CT_CellAlignment */
			case '<alignment': case '<alignment/>': break;

			/* 18.8.33 protection CT_CellProtection */
			case '<protection': case '</protection>': case '<protection/>': break;

			case '<extLst': case '</extLst>': break;
			case '<ext': break;
			default: if(opts.WTF) throw 'unrecognized ' + y[0] + ' in cellXfs';
		}
	});
}

function write_cellXfs(cellXfs) {
	var o = [];
	o[o.length] = (writextag('cellXfs',null));
	cellXfs.forEach(function(c) { o[o.length] = (writextag('xf', null, c)); });
	o[o.length] = ("</cellXfs>");
	if(o.length === 2) return "";
	o[0] = writextag('cellXfs',null, {count:o.length-2}).replace("/>",">");
	return o.join("");
}

/* 18.8 Styles CT_Stylesheet*/
var parse_sty_xml= (function make_pstyx() {
var numFmtRegex = /<numFmts([^>]*)>.*<\/numFmts>/;
var cellXfRegex = /<cellXfs([^>]*)>.*<\/cellXfs>/;
var fillsRegex = /<fills([^>]*)>.*<\/fills>/;

return function parse_sty_xml(data, opts) {
	/* 18.8.39 styleSheet CT_Stylesheet */
	var t;

	/* numFmts CT_NumFmts ? */
	if((t=data.match(numFmtRegex))) parse_numFmts(t, opts);

	/* fonts CT_Fonts ? */
	/*if((t=data.match(/<fonts([^>]*)>.*<\/fonts>/))) parse_fonts(t, opts);*/

	/* fills CT_Fills */
	if((t=data.match(fillsRegex))) parse_fills(t, opts);

	/* borders CT_Borders ? */
	/* cellStyleXfs CT_CellStyleXfs ? */

	/* cellXfs CT_CellXfs ? */
	if((t=data.match(cellXfRegex))) parse_cellXfs(t, opts);

	/* dxfs CT_Dxfs ? */
	/* tableStyles CT_TableStyles ? */
	/* colors CT_Colors ? */
	/* extLst CT_ExtensionList ? */

	return styles;
};
})();

var STYLES_XML_ROOT = writextag('styleSheet', null, {
	'xmlns': XMLNS.main[0],
	'xmlns:vt': XMLNS.vt
});

RELS.STY = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

function write_sty_xml(wb, opts) {
	var o = [XML_HEADER, STYLES_XML_ROOT], w;
	if((w = write_numFmts(wb.SSF)) != null) o[o.length] = w;
	if((w = write_fonts(opts.fonts))) o[o.length] = (w);
	if((w = write_fills(opts.fills))) o[o.length] = (w);
	if((w = write_borders(opts.borders))) o[o.length] = (w);
	o[o.length] = ('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>');
	if((w = write_cellXfs(opts.cellXfs))) o[o.length] = (w);
	o[o.length] = ('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>');
	o[o.length] = ('<dxfs count="0"/>');
	o[o.length] = ('<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleMedium4"/>');

	if(o.length>2){ o[o.length] = ('</styleSheet>'); o[1]=o[1].replace("/>",">"); }
	return o.join("");
}
