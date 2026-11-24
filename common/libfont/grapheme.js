/*
 * (c) Copyright Ascensio System SIA 2010-2024
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

"use strict";

(function(window)
{
	const FONTSIZE        = AscFonts.MEASURE_FONTSIZE || 576;
	const STRING_MAX_LEN  = AscFonts.GRAPHEME_STRING_MAX_LEN || 1024;
	const COEF            = AscFonts.GRAPHEME_COEF|| 25.4 / 72 / 64 / FONTSIZE;
	const GRAPHEME_BUFFER = new Array(STRING_MAX_LEN);
	let   GRAPHEME_LEN    = 0;
	const GRAPHEMES_CACHE = {};                                 // FontId + [GID] -> GraphemeId
	const GRAPHEMES       = [[0, 0, 0, 0, 0, 0, 0, 0, [0x20]]]; // GraphemeId -> Grapheme
	let   GRAPHEME_INDEX  = 0;
	const NO_GRAPHEME     = 0;

	function InitGrapheme(nFontId, nFontStyle)
	{
		GRAPHEME_LEN = 3;
		GRAPHEME_BUFFER[0] = nFontId << 8 | nFontStyle;
		GRAPHEME_BUFFER[1] = 0;
		GRAPHEME_BUFFER[2] = 0;
	}
	function AddGlyphToGrapheme(nGID, nAdvanceX, nAdvanceY, nOffsetX, nOffsetY)
	{
		GRAPHEME_BUFFER[GRAPHEME_LEN++] = nGID;
		GRAPHEME_BUFFER[GRAPHEME_LEN++] = nAdvanceX;
		GRAPHEME_BUFFER[GRAPHEME_LEN++] = nAdvanceY;
		GRAPHEME_BUFFER[GRAPHEME_LEN++] = nOffsetX;
		GRAPHEME_BUFFER[GRAPHEME_LEN++] = nOffsetY;
		GRAPHEME_LEN++; // CodePoints

		GRAPHEME_BUFFER[1] += nAdvanceX;
		GRAPHEME_BUFFER[2] += 1;
	}
	function DrawGrapheme(nGraphemeId, oContext, nX, nY, nFontSize, coeff)
	{
		let oGrapheme = GRAPHEMES[nGraphemeId];
		if (!oGrapheme)
			return;
		
		if (undefined === coeff)
			coeff = 1;

		let nFontId = oGrapheme[0] >> 8;
		let nStyle  = oGrapheme[0] & 0xF;

		let sFontName = AscCommon.FontNameMap.GetName(nFontId);
		oContext.SetFontInternal(sFontName, nFontSize, nStyle);
		
		let nKoef = COEF * nFontSize * coeff;

		let oColorGlyphs = GetColorGlyphDrawer(oGrapheme[0], sFontName, nStyle, oContext);
		if (oColorGlyphs)
		{
			oColorGlyphs.Begin();
			let nPos = 3;
			for (let nIndex = 0, nCount = oGrapheme[2]; nIndex < nCount; ++nIndex)
			{
				let nGID          = oGrapheme[nPos++];
				let nAdvanceX     = oGrapheme[nPos++];
				let nAdvanceY     = oGrapheme[nPos++];
				let nOffsetX      = oGrapheme[nPos++];
				let nOffsetY      = oGrapheme[nPos++];
				let arrCodePoints = oGrapheme[nPos++];

				oColorGlyphs.Draw(nGID, nX + nOffsetX * nKoef, nY - nOffsetY * nKoef, arrCodePoints);
				nX += nAdvanceX * nKoef;
				nY += nAdvanceY * nKoef;
			}
			oColorGlyphs.End();
		}
		else if (1 === oGrapheme[2])
		{
			oContext.tg(oGrapheme[3], nX + oGrapheme[6] * nKoef, nY - oGrapheme[7] * nKoef, oGrapheme[8]);
		}
		else
		{
			let nKoef = COEF * nFontSize * coeff;
			let nPos  = 3;
			for (let nIndex = 0, nCount = oGrapheme[2]; nIndex < nCount; ++nIndex)
			{
				let nGID          = oGrapheme[nPos++];
				let nAdvanceX     = oGrapheme[nPos++];
				let nAdvanceY     = oGrapheme[nPos++];
				let nOffsetX      = oGrapheme[nPos++];
				let nOffsetY      = oGrapheme[nPos++];
				let arrCodePoints = oGrapheme[nPos++];
				
				oContext.tg(nGID, nX + nOffsetX * nKoef, nY - nOffsetY * nKoef, arrCodePoints);
				nX += nAdvanceX * nKoef;
				nY += nAdvanceY * nKoef;
			}
		}
	}
	//------------------------------------------------------------------------------------------------------------------
	// Colour glyphs. PowerPoint draws the glyphs of a colour font (COLR/CPAL, e.g. Segoe UI Emoji) in their palette
	// colours; the font engine and the PDF writer only know the monochrome outline. For slides, a glyph that has a
	// COLR v0 layer list is drawn as its layers, bottom to top, each filled with its palette 0 colour (or with the
	// text colour for palette index 0xFFFF).
	//------------------------------------------------------------------------------------------------------------------
	const COLOR_GLYPH_TABLES = {}; // font file id + face index -> CColorGlyphTable or null
	const COLOR_GLYPH_FONTS  = {}; // grapheme font key (FontNameMap id << 8 | style) -> CColorGlyphTable or null
	const NO_CODEPOINTS      = [];

	function CColorGlyphTable(arrBase, arrLayers, arrPalette)
	{
		this.Base    = arrBase;    // [gid, firstLayerIndex, numLayers] * n, sorted by gid
		this.Layers  = arrLayers;  // [gid, paletteIndex] * n
		this.Palette = arrPalette; // [r, g, b, a] * n, palette 0
	}
	CColorGlyphTable.prototype.FindBase = function(nGID)
	{
		let nL = 0, nR = this.Base.length / 3 - 1;
		while (nL <= nR)
		{
			let nM   = (nL + nR) >> 1;
			let nCur = this.Base[3 * nM];
			if (nCur === nGID)
				return nM;
			else if (nCur < nGID)
				nL = nM + 1;
			else
				nR = nM - 1;
		}
		return -1;
	};
	function ReadColorGlyphTable(data, size, faceIndex)
	{
		if (!data || "string" === typeof data || !(size > 12))
			return null;

		let u8  = function(pos) { return data[pos] & 0xFF; };
		let u16 = function(pos) { return (u8(pos) << 8) | u8(pos + 1); };
		let u32 = function(pos) { return ((u8(pos) << 24) >>> 0) + (u8(pos + 1) << 16) + (u8(pos + 2) << 8) + u8(pos + 3); };

		let offset = 0;
		if (0x74746366 === u32(0)) // 'ttcf'
		{
			if (faceIndex >= u32(8))
				return null;
			offset = u32(12 + 4 * faceIndex);
		}

		let tag = u32(offset);
		if (0x00010000 !== tag && 0x4F54544F !== tag && 0x74727565 !== tag) // 1.0, 'OTTO', 'true'
			return null;

		let colr = null, cpal = null;
		for (let i = 0, numTables = u16(offset + 4); i < numTables; ++i)
		{
			let rec = offset + 12 + 16 * i;
			if (rec + 16 > size)
				return null;
			let recTag = u32(rec);
			if (0x434F4C52 === recTag) // 'COLR'
				colr = {offset : u32(rec + 8), length : u32(rec + 12)};
			else if (0x4350414C === recTag) // 'CPAL'
				cpal = {offset : u32(rec + 8), length : u32(rec + 12)};
		}

		if (!colr || !cpal || colr.length < 14 || cpal.length < 14 || colr.offset + 14 > size || cpal.offset + 14 > size)
			return null;

		// COLR v0 header (v1 keeps the v0 fields and records first)
		let c = colr.offset;
		if (u16(c) > 1)
			return null;
		let numBase     = u16(c + 2);
		let baseOffset  = c + u32(c + 4);
		let layerOffset = c + u32(c + 8);
		let numLayers   = u16(c + 12);
		if (!numBase || !numLayers || baseOffset + 6 * numBase > size || layerOffset + 4 * numLayers > size)
			return null;

		// CPAL, palette 0; colour records are B, G, R, A
		let p            = cpal.offset;
		let numEntries   = u16(p + 2);
		let numPalettes  = u16(p + 4);
		let numRecords   = u16(p + 6);
		let recordsStart = p + u32(p + 8);
		if (!numEntries || !numPalettes || p + 14 > size)
			return null;
		let firstRecord = u16(p + 12);
		if (firstRecord + numEntries > numRecords || recordsStart + 4 * (firstRecord + numEntries) > size)
			return null;

		let arrBase = new Uint16Array(3 * numBase);
		for (let i = 0; i < numBase; ++i)
		{
			let pos = baseOffset + 6 * i;
			arrBase[3 * i]     = u16(pos);
			arrBase[3 * i + 1] = u16(pos + 2);
			arrBase[3 * i + 2] = u16(pos + 4);
		}

		let arrLayers = new Uint16Array(2 * numLayers);
		for (let i = 0; i < numLayers; ++i)
		{
			let pos = layerOffset + 4 * i;
			arrLayers[2 * i]     = u16(pos);
			arrLayers[2 * i + 1] = u16(pos + 2);
		}

		let arrPalette = new Uint8Array(4 * numEntries);
		for (let i = 0; i < numEntries; ++i)
		{
			let pos = recordsStart + 4 * (firstRecord + i);
			arrPalette[4 * i]     = u8(pos + 2);
			arrPalette[4 * i + 1] = u8(pos + 1);
			arrPalette[4 * i + 2] = u8(pos);
			arrPalette[4 * i + 3] = u8(pos + 3);
		}

		return new CColorGlyphTable(arrBase, arrLayers, arrPalette);
	}
	function GetColorGlyphTable(nFontKey, sFontName, nStyle)
	{
		if (undefined !== COLOR_GLYPH_FONTS[nFontKey])
			return COLOR_GLYPH_FONTS[nFontKey];

		let oInfo = AscFonts.g_fontApplication ? AscFonts.g_fontApplication.GetFontInfo(sFontName, nStyle) : null;
		let oId   = oInfo && AscCommon.g_font_loader ? oInfo.GetFontID(AscCommon.g_font_loader, nStyle) : null;
		if (!oId || !oId.file)
			return (COLOR_GLYPH_FONTS[nFontKey] = null);

		let nStreamIndex = oId.file.stream_index;
		if (undefined === nStreamIndex || null === nStreamIndex || nStreamIndex < 0)
			return null; // not loaded yet: ask again next time

		let nFaceIndex = oId.faceIndex ? oId.faceIndex : 0;
		let sKey       = oId.id + "_" + nFaceIndex;
		if (undefined === COLOR_GLYPH_TABLES[sKey])
		{
			let oTable = null;
			try
			{
				let oStream = AscFonts["getFontStream"](nStreamIndex);
				oTable = oStream ? ReadColorGlyphTable(oStream["data"], oStream["size"], nFaceIndex) : null;
			}
			catch (e)
			{
				oTable = null;
			}
			COLOR_GLYPH_TABLES[sKey] = oTable;
		}

		return (COLOR_GLYPH_FONTS[nFontKey] = COLOR_GLYPH_TABLES[sKey]);
	}

	/**
	 * Draws the glyphs of one grapheme as COLR layers on a drawing context with a solid text brush.
	 * @constructor
	 */
	function CColorGlyphDrawer()
	{
		this.Table   = null;
		this.Context = null;
		this.R       = 0;
		this.G       = 0;
		this.B       = 0;
		this.A       = 255;
		this.Changed = false;
	}
	CColorGlyphDrawer.prototype.Init = function(oTable, oContext, oColor)
	{
		this.Table   = oTable;
		this.Context = oContext;
		this.R       = oColor.R;
		this.G       = oColor.G;
		this.B       = oColor.B;
		this.A       = undefined === oColor.A ? 255 : oColor.A;
		this.Changed = false;
		return this;
	};
	CColorGlyphDrawer.prototype.Begin = function()
	{
		this.Changed = false;
	};
	CColorGlyphDrawer.prototype.Draw = function(nGID, nX, nY, arrCodePoints)
	{
		let oTable = this.Table;
		let nBase  = oTable.FindBase(nGID);
		if (-1 === nBase)
		{
			if (this.Changed)
			{
				this.Context.b_color1(this.R, this.G, this.B, this.A);
				this.Changed = false;
			}
			this.Context.tg(nGID, nX, nY, arrCodePoints);
			return;
		}

		let nFirst      = oTable.Base[3 * nBase + 1];
		let nCount      = oTable.Base[3 * nBase + 2];
		let nPalCount   = oTable.Palette.length >> 2;
		let nLayerCount = oTable.Layers.length >> 1;
		for (let nIndex = 0; nIndex < nCount && nFirst + nIndex < nLayerCount; ++nIndex)
		{
			let nLayer   = nFirst + nIndex;
			let nPalette = oTable.Layers[2 * nLayer + 1];
			if (nPalette < nPalCount)
			{
				let p = 4 * nPalette;
				let a = (oTable.Palette[p + 3] * this.A / 255 + 0.5) | 0;
				this.Context.b_color1(oTable.Palette[p], oTable.Palette[p + 1], oTable.Palette[p + 2], a);
				this.Changed = true;
			}
			else if (this.Changed)
			{
				this.Context.b_color1(this.R, this.G, this.B, this.A);
				this.Changed = false;
			}

			// The code points go with the first layer only, so that the text is extracted once
			this.Context.tg(oTable.Layers[2 * nLayer], nX, nY, 0 === nIndex ? arrCodePoints : NO_CODEPOINTS);
		}
	};
	CColorGlyphDrawer.prototype.End = function()
	{
		if (this.Changed)
			this.Context.b_color1(this.R, this.G, this.B, this.A);

		this.Changed = false;
		this.Table   = null;
		this.Context = null;
	};
	const COLOR_GLYPH_DRAWER = new CColorGlyphDrawer();

	/**
	 * @returns {?CColorGlyphDrawer} a drawer if the grapheme's font has colour glyphs and the context draws text with
	 * a solid brush we can read back (canvas or PDF metafile), on slides only
	 */
	function GetColorGlyphDrawer(nFontKey, sFontName, nStyle, oContext)
	{
		if (!Asc.editor || !Asc.editor.isPresentationEditor)
			return null;

		let oTable = GetColorGlyphTable(nFontKey, sFontName, nStyle);
		if (!oTable)
			return null;

		let oBrush = null;
		if (AscCommon.CDocumentRenderer && oContext instanceof AscCommon.CDocumentRenderer)
		{
			let oPage = oContext.m_lPagesCount > 0 ? oContext.m_arrayPages[oContext.m_lPagesCount - 1] : null;
			if (oPage && 0 === oPage.BrushType) // MetaBrushType.Solid
				oBrush = oPage.m_oBrush;
		}
		else if (AscCommon.CGraphics && oContext instanceof AscCommon.CGraphics)
		{
			oBrush = oContext.m_oBrush;
		}

		let oColor = oBrush ? oBrush.Color1 : null;
		if (!oColor || !(oColor.R >= 0) || !(oColor.G >= 0) || !(oColor.B >= 0))
			return null;

		return COLOR_GLYPH_DRAWER.Init(oTable, oContext, oColor);
	}
	function CompareGraphemes(g)
	{
		if (g.length !== GRAPHEME_LEN)
			return false;

		for (let nPos = 0; nPos < GRAPHEME_LEN; ++nPos)
		{
			if (g[nPos] !== GRAPHEME_BUFFER[nPos])
				return false;
		}

		return true;
	}
	function FillCodePoints(codePoints)
	{
		// Пишем по схеме: распределяем везде по 1, если есть лишние, то все они уходят в первый глиф, если
		// последним глифам не хватает, тогда им проставляем пробелы

		let nCount = GRAPHEME_BUFFER[2];
		if (nCount <= 0)
			return;

		let nCodePointPos = 0;
		let nCodePointsCount = codePoints ? codePoints.getCount() : 0;

		if (nCount > 1 && FillCodePointsByCmap(codePoints, nCount, nCodePointsCount))
			return;

		// A right-to-left cluster comes from the shaper in visual order, its marks before the base
		// glyph, so the code points go to its glyphs from the last one back: the base glyph gets the
		// base letter. (PDF output picks glyphs by these code points, so a mark glyph holding the
		// letter drew a wrong letter elsewhere.)
		let isRtl = !!(codePoints && codePoints.isRtl);
		for (let nGlyph = 0; nGlyph < nCount; ++nGlyph)
		{
			let nPos = 8 + 6 * (isRtl ? nCount - 1 - nGlyph : nGlyph);
			if (nCodePointPos >= nCodePointsCount)
			{
				GRAPHEME_BUFFER[nPos] = [0x20];
			}
			else if (0 === nGlyph)
			{
				GRAPHEME_BUFFER[nPos] = [];
				let nStoreCount = Math.max(1, nCodePointsCount - nCount + 1);
				for (;nCodePointPos < nStoreCount; ++nCodePointPos)
				{
					GRAPHEME_BUFFER[nPos].push(codePoints.get(nCodePointPos));
				}
			}
			else
			{
				GRAPHEME_BUFFER[nPos] = [codePoints.get(nCodePointPos++)];
			}
		}
	}
	/**
	 * Gives each glyph of a multi-glyph cluster the code point whose cmap glyph it is (a mark glyph
	 * gets its mark), and the code points no glyph stands for alone (a base letter composed with a
	 * mark, say bet + dagesh) to the remaining glyphs in logical order, the first of them taking
	 * any extra. HarfBuzz reorders marks (canonical order, visual order for right-to-left text) and
	 * composes them, so glyphs and code points do not line up by position. PDF output picks glyphs
	 * by these code points.
	 * @returns {boolean} false when the cmap glyphs of the code points are not known
	 */
	function FillCodePointsByCmap(codePoints, nCount, nCodePointsCount)
	{
		if (!codePoints || !codePoints.hasGids || !codePoints.hasGids() || nCodePointsCount <= 0)
			return false;

		let isRtl    = !!codePoints.isRtl;
		let assigned = new Array(nCount);
		let used     = new Array(nCodePointsCount);
		for (let nGlyph = 0; nGlyph < nCount; ++nGlyph)
		{
			let nGID = GRAPHEME_BUFFER[3 + 6 * nGlyph];
			for (let nCodePoint = 0; nCodePoint < nCodePointsCount; ++nCodePoint)
			{
				if (!used[nCodePoint] && codePoints.getGid(nCodePoint) === nGID)
				{
					assigned[nGlyph]  = [codePoints.get(nCodePoint)];
					used[nCodePoint] = true;
					break;
				}
			}
		}

		let rest = [];
		for (let nCodePoint = 0; nCodePoint < nCodePointsCount; ++nCodePoint)
		{
			if (!used[nCodePoint])
				rest.push(codePoints.get(nCodePoint));
		}

		// Glyphs without a code point of their own, in logical order
		let free = [];
		for (let nIndex = 0; nIndex < nCount; ++nIndex)
		{
			let nGlyph = isRtl ? nCount - 1 - nIndex : nIndex;
			if (!assigned[nGlyph])
				free.push(nGlyph);
		}

		if (free.length > 0)
		{
			let nExtra = Math.max(0, rest.length - free.length);
			let nRest  = 0;
			for (let nIndex = 0; nIndex < free.length; ++nIndex)
			{
				let nTake = 0 === nIndex ? 1 + nExtra : 1;
				let arr   = rest.slice(nRest, nRest + nTake);
				nRest += nTake;
				assigned[free[nIndex]] = arr.length ? arr : [0x20];
			}
		}
		else if (rest.length > 0)
		{
			// Every glyph is a cmap glyph and code points remain: keep them with the first glyph
			let nFirst = isRtl ? nCount - 1 : 0;
			assigned[nFirst] = rest.concat(assigned[nFirst]);
		}

		for (let nGlyph = 0; nGlyph < nCount; ++nGlyph)
			GRAPHEME_BUFFER[8 + 6 * nGlyph] = assigned[nGlyph];

		return true;
	}
	function GetGraphemeIndex(codePoints)
	{
		FillCodePoints(codePoints);

		let arrBuffer = new Array(GRAPHEME_LEN);
		for (let nIndex = 0; nIndex < GRAPHEME_LEN; ++nIndex)
			arrBuffer[nIndex] = GRAPHEME_BUFFER[nIndex];

		GRAPHEMES[++GRAPHEME_INDEX] = arrBuffer;
		return GRAPHEME_INDEX;
	}
	function GetGrapheme(codePoints)
	{
		let nFontId = GRAPHEME_BUFFER[0];
		if (!GRAPHEMES_CACHE[nFontId])
			GRAPHEMES_CACHE[nFontId] = {};

		let result = GRAPHEMES_CACHE[nFontId];

		let nPos = 3;
		while (nPos < GRAPHEME_LEN)
		{
			let nGID = GRAPHEME_BUFFER[nPos];
			nPos += 6;

			if (!result[nGID])
				result[nGID] = {};

			result = result[nGID];
		}

		// TODO: Для скорости проверку совпадения отключили (всегда совпадает)
		if (!result.Grapheme)
			result.Grapheme = GetGraphemeIndex(codePoints);
		// else if (!CompareGraphemes(result.Buffer))
		// 	return GetGraphemeIndex(codePoints);

		return result.Grapheme;
	}
	function GetGraphemeWidth(nGraphemeId)
	{
		let oGrapheme = GRAPHEMES[nGraphemeId];
		if (!oGrapheme)
			return 0;

		return oGrapheme[1] * COEF;
	}
	function GetGraphemeBBox(graphemeId, fontSize, dpi)
	{
		let grapheme = GRAPHEMES[graphemeId];
		if (!grapheme || grapheme[2] <= 0)
			return {minX : 0, maxX : 0, minY : 0, maxY : 0};
		
		let measurer = AscCommon.g_oTextMeasurer;
		
		if (undefined === dpi)
			dpi = 72;
		
		let fontId    = grapheme[0] >> 8;
		let fontStyle = grapheme[0] & 0xF;
		
		let fontName = AscCommon.FontNameMap.GetName(fontId);
		measurer.SetFontInternal(fontName, fontSize, fontStyle, dpi);
		
		let bbox = measurer.GetBBox(grapheme[3]);
		
		let minX = bbox.fMinX;
		let maxX = bbox.fMaxX;
		let minY = bbox.fMinY;
		let maxY = bbox.fMaxY;
		
		if (grapheme[2] > 1)
		{
			let x = 0;
			let y = 0;
			
			let grCoeff = COEF * fontSize * dpi / 25.4;
			for (let pos = 3, index = 0, count = grapheme[2]; index < count; ++index)
			{
				let bbox = measurer.GetBBox(grapheme[pos++]);
				
				minX = Math.min(minX, x + bbox.fMinX);
				maxX = Math.max(maxX, x + bbox.fMaxX);
				minY = Math.min(minY, y + bbox.fMinY);
				maxY = Math.max(maxY, y + bbox.fMaxY);
				
				x += grapheme[pos++] * grCoeff; // advanceX
				y += grapheme[pos++] * grCoeff; // advanceY
				
				pos += 3;
			}
		}
		
		return {
			minX : minX,
			maxX : maxX,
			minY : minY,
			maxY : maxY
		};
	}
	function GetGraphemeFontId(graphemeId)
	{
		let oGrapheme = GRAPHEMES[graphemeId];
		if (!oGrapheme)
			return null;
		
		return oGrapheme[0];
	}
	/**
	 * The glyph id of a grapheme made of exactly one glyph, otherwise -1.
	 */
	function GetGraphemeSingleGid(graphemeId)
	{
		let grapheme = GRAPHEMES[graphemeId];
		if (!grapheme || 1 !== grapheme[2])
			return -1;

		return grapheme[3];
	}
	function GetFontNameByFontId(fontId)
	{
		return AscCommon.FontNameMap.GetName((fontId | 0) >> 8);
	}
	function GetFontStyleByFontId(fontId)
	{
		return ((fontId | 0) & 0xF);
	}
	function GetGraphemeCodePoints(graphemeId)
	{
		let g = GRAPHEMES[graphemeId];
		if (!g)
			return [];
		
		if (1 === g[2])
		{
			return g[8];
		}
		else
		{
			let result = [];
			for (let i = 0, pos = 8, count = g[2]; i < count; ++i, pos += 6)
			{
				result = result.concat(g[pos]);
			}
			return result;
		}
		
	}
	//--------------------------------------------------------export----------------------------------------------------
	window['AscFonts'] = window['AscFonts'] || {};
	window['AscFonts'].NO_GRAPHEME           = NO_GRAPHEME;
	window['AscFonts'].InitGrapheme          = InitGrapheme;
	window['AscFonts'].DrawGrapheme          = DrawGrapheme;
	window['AscFonts'].CompareGraphemes      = CompareGraphemes;
	window['AscFonts'].AddGlyphToGrapheme    = AddGlyphToGrapheme;
	window['AscFonts'].GetGrapheme           = GetGrapheme;
	window['AscFonts'].GetGraphemeWidth      = GetGraphemeWidth;
	window['AscFonts'].GetGraphemeBBox       = GetGraphemeBBox;
	window['AscFonts'].GetGraphemeFontId     = GetGraphemeFontId;
	window['AscFonts'].GetGraphemeSingleGid  = GetGraphemeSingleGid;
	window['AscFonts'].GetFontNameByFontId   = GetFontNameByFontId;
	window['AscFonts'].GetFontStyleByFontId  = GetFontStyleByFontId;
	window['AscFonts'].GetGraphemeCodePoints = GetGraphemeCodePoints;

})(window);
