var strs = {}; // shared strings
var _ssfopts = {}; // spreadsheet formatting options

RELS.WS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

function get_sst_id(sst, str) {
	for(var i = 0, len = sst.length; i < len; ++i) if(sst[i].t === str) { sst.Count ++; return i; }
	sst[len] = {t:str}; sst.Count ++; sst.Unique ++; return len;
}

function get_cell_style(styles, cell, opts) {
	var z = opts.revssf[cell.z != null ? cell.z : "General"];
	var f = 0;
	var b = 0;
	if (cell.s) {
		if (cell.s.fill) {
			var fs = JSON.stringify(cell.s.fill);
			if (!opts.fills) {
				opts.fills = [fs];
				f = 1;
			}
			else {
				f = opts.fills.indexOf(fs) + 1;
				if (f === 0) {
					opts.fills.push(fs);
					f = opts.fills.length;
				}
			}
		}
		if (cell.s.border) {
			var bs = JSON.stringify(cell.s.border);
			if (!opts.borders) {
				opts.borders = [bs];
				b = 1;
			}
			else {
				b = opts.borders.indexOf(bs) + 1;
				if (b === 0) {
					opts.borders.push(bs);
					b = opts.borders.length;
				}
			}
		}
	}
	for(var i = 0, len = styles.length; i != len; ++i) {
		if (styles[i].numFmtId === z &&
			styles[i].fillId === f &&
			styles[i].borderId === b)
			return i;
	}
	styles[len] = {
		numFmtId:z,
		fontId:0,
		fillId:f,
		borderId:b,
		xfId:0,
		applyNumberFormat:1
	};
	return len;
}

function safe_format(p, fmtid, fillid, opts) {
	try {
		if(p.t === 'e') p.w = p.w || BErr[p.v];
		else if(fmtid === 0) {
			if(p.t === 'n') {
				if((p.v|0) === p.v) p.w = SSF._general_int(p.v,_ssfopts);
				else p.w = SSF._general_num(p.v,_ssfopts);
			}
			else if(p.t === 'd') {
				var dd = datenum(p.v);
				if((dd|0) === dd) p.w = SSF._general_int(dd,_ssfopts);
				else p.w = SSF._general_num(dd,_ssfopts);
			}
			else if(p.v === undefined) return "";
			else p.w = SSF._general(p.v,_ssfopts);
		}
		else if(p.t === 'd') p.w = SSF.format(fmtid,datenum(p.v),_ssfopts);
		else p.w = SSF.format(fmtid,p.v,_ssfopts);
		if(opts.cellNF) p.z = SSF._table[fmtid];
	} catch(e) { if(opts.WTF) throw e; }
	if(fillid) try {
		p.s = styles.Fills[fillid];
		if (p.s.fgColor && p.s.fgColor.theme) {
			p.s.fgColor.rgb = rgb_tint(themes.themeElements.clrScheme[p.s.fgColor.theme].rgb, p.s.fgColor.tint || 0);
			if(opts.WTF) p.s.fgColor.raw_rgb = themes.themeElements.clrScheme[p.s.fgColor.theme].rgb;
		}
		if (p.s.bgColor && p.s.bgColor.theme) {
			p.s.bgColor.rgb = rgb_tint(themes.themeElements.clrScheme[p.s.bgColor.theme].rgb, p.s.bgColor.tint || 0);
			if(opts.WTF) p.s.bgColor.raw_rgb = themes.themeElements.clrScheme[p.s.bgColor.theme].rgb;
		}
	} catch(e) { if(opts.WTF) throw e; }
}
