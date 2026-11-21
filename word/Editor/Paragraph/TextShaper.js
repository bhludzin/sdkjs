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
	const CODEPOINT_TYPE = {
		BASE              : 0,
		LIGATURE          : 1,
		LIGATURE_CONTINUE : 2,
		COMBINING_MARK    : 3
	};

	/**
	 *
	 * @constructor
	 */
	function CParagraphTextShaper()
	{
		AscFonts.CTextShaper.call(this);

		this.Parent    = null;
		this.Paragraph = null;
		this.TextPr    = null;
		this.Temporary = false;
		this.Ligatures = Asc.LigaturesType.None;
		this.Spacing   = 0;
		this.AscFont   = false; // Специальный случай, когда используемый шрифт ASCW3, а не тот, что задан в настройках
		
		this.MaskSymbol = null; // Символ, который используется для маскирования текста в полях ввода
	}
	CParagraphTextShaper.prototype = Object.create(AscFonts.CTextShaper.prototype);
	CParagraphTextShaper.prototype.constructor = CParagraphTextShaper;

	CParagraphTextShaper.prototype.Init = function(isTemporary)
	{
		this.Parent    = null;
		this.Paragraph = null;
		this.TextPr    = null;
		this.Temporary = isTemporary;
		this.Ligatures = Asc.LigaturesType.None;
		this.Spacing   = 0;
		this.AscFont   = false;
		
		this.ClearBuffer();
	};
	CParagraphTextShaper.prototype.GetCodePoint = function(oItem)
	{
		if (null !== this.MaskSymbol)
			return this.MaskSymbol;
		
		let nCodePoint = oItem.GetCodePoint();

		if (this.TextPr && (this.TextPr.Caps || this.TextPr.SmallCaps))
			nCodePoint = (String.fromCharCode(nCodePoint).toUpperCase()).charCodeAt(0);

		return nCodePoint;
	};
	CParagraphTextShaper.prototype.GetFontInfo = function(nFontSlot)
	{
		if (!this.TextPr)
			return AscFonts.DEFAULT_TEXTFONTINFO;

		let oFontInfo = this.TextPr.GetFontInfo(nFontSlot);

		if (this.AscFont)
			oFontInfo.Name = "ASCW3";

		return oFontInfo;
	};
	CParagraphTextShaper.prototype.GetFontSlot = function(nUnicode)
	{
		let oTextPr = this.TextPr;
		if (!oTextPr)
			return AscWord.fontslot_None;

		return AscWord.GetFontSlotByTextPr(nUnicode, oTextPr);
	};
	CParagraphTextShaper.prototype.GetLigaturesType = function(textScript)
	{
		// bug-73560 
		let result = this.Ligatures;
		if (AscFonts.HB_SCRIPT.HB_SCRIPT_ARABIC === textScript)
			result |= Asc.LigaturesType.Standard;
		
		// Pair kerning from the run's kern threshold (DrawingML a:rPr/@kern)
		if (this.TextPr && this.TextPr.IsKerned())
			result |= AscFonts.HB_FEATURE.HB_FEATURE_TYPE_KERNING;

		return result;
	};
	CParagraphTextShaper.prototype.Shape = function(oParagraph)
	{
		this.Init(false);
		let oThis = this;
		oParagraph.CheckRunContent(function(oRun, nStartPos, nEndPos)
		{
			oThis.HandleRun(oRun, nStartPos, nEndPos);
		});
		this.FlushWord();
	};
	CParagraphTextShaper.prototype.ShapeRange = function(oParagraph, oStart, oEnd, isTemporary)
	{
		this.Init(isTemporary);
		let oThis = this;
		oParagraph.CheckRunContent(function(oRun, nStartPos, nEndPos)
		{
			oThis.HandleRun(oRun, nStartPos, nEndPos);
		}, oStart, oEnd);
		this.FlushWord();
	};
	CParagraphTextShaper.prototype.ShapeRun = function(run)
	{
		this.Init(false);
		this.HandleRun(run, 0, run.GetElementsCount());
		this.FlushWord();
	};
	CParagraphTextShaper.prototype.HandleRun = function(oRun, nStartPos, nEndPos)
	{
		this.private_CheckRun(oRun);

		for (let nPos = nStartPos; nPos < nEndPos; ++nPos)
		{
			let oItem = oRun.GetElement(nPos);
			if (oItem.IsPdfText())
			{
				this.FlushWord();
				this.private_HandlePdfText(oItem);
			}
			else if (!oItem.IsText())
			{
				this.FlushWord();
				if (oItem.IsSpace())
					this.private_HandleSpace(oItem);
			}
			else if (oItem.IsNBSP())
			{
				this.FlushWord();
				this.private_HandleNBSP(oItem);
			}
			else if (oItem.IsDigit() && this.private_IsReplaceToHindiDigits())
			{
				this.FlushWord();
				this.private_HandleHindiDigit(oItem);
			}
			else
			{
				this.AppendToString(oItem);

				if (oItem.IsSpaceAfter())
					this.FlushWord();
			}
		}
	};
	CParagraphTextShaper.prototype.FlushWord = function()
	{
		// The glyph advances PowerPoint rounds are measured here, before shaping, because the
		// font engine cannot be called while the shaped glyphs are read (see getGlyphAdvance).
		if (this.Buffer.length > 0 && -1 !== this.FontId && this.FontId && this.FontId.m_pFaceInfo
			&& this.private_GetLayoutFontSize(1) > 0)
		{
			let codePoints = [];
			for (let index = 0, count = this.Buffer.length; index < count; ++index)
				codePoints.push(this.GetCodePoint(this.Buffer[index]));

			let fontInfo = this.GetFontInfo(this.FontSlot);
			prefetchGlyphAdvances(this.FontId.m_pFaceInfo.family_name, fontInfo.Style, codePoints);
		}

		AscFonts.CTextShaper.prototype.FlushWord.call(this);
	};
	CParagraphTextShaper.prototype.FlushGrapheme = function(nGrapheme, nWidth, nCodePointsCount, isLigature)
	{
		if (nCodePointsCount <= 0)
			return;
		
		let curIndex = 0;
		
		if (this.IsRtlDirection())
		{
			if (this.BufferIndex - nCodePointsCount < 0)
				return;
			
			this.BufferIndex -= nCodePointsCount;
			curIndex = this.BufferIndex;
		}
		else
		{
			if (this.BufferIndex + nCodePointsCount - 1 >= this.Buffer.length)
				return;
			
			curIndex = this.BufferIndex;
			this.BufferIndex += nCodePointsCount;
		}
		
		if (1 === nCodePointsCount)
			nWidth = this.private_GetLayoutWidth(nGrapheme, nWidth);

		let _nWidth = (nWidth + (this.Spacing / this.FontSize)) / nCodePointsCount;
		if (1 === nCodePointsCount)
		{
			this.private_HandleItem(this.Buffer[curIndex], nGrapheme, _nWidth, this.FontSize, this.FontSlot, CODEPOINT_TYPE.BASE);
		}
		else
		{
			if (this.IsRtlDirection())
			{
				this.private_HandleItem(this.Buffer[curIndex], AscFonts.NO_GRAPHEME, _nWidth, this.FontSize, this.FontSlot, isLigature ? CODEPOINT_TYPE.LIGATURE : CODEPOINT_TYPE.BASE);
				this.private_HandleItem(this.Buffer[curIndex + nCodePointsCount - 1], nGrapheme, _nWidth, this.FontSize, this.FontSlot, isLigature ? CODEPOINT_TYPE.LIGATURE_CONTINUE : CODEPOINT_TYPE.COMBINING_MARK);
			}
			else
			{
				this.private_HandleItem(this.Buffer[curIndex], nGrapheme, _nWidth, this.FontSize, this.FontSlot, isLigature ? CODEPOINT_TYPE.LIGATURE : CODEPOINT_TYPE.BASE);
				this.private_HandleItem(this.Buffer[curIndex + nCodePointsCount - 1], AscFonts.NO_GRAPHEME, _nWidth, this.FontSize, this.FontSlot, isLigature ? CODEPOINT_TYPE.LIGATURE_CONTINUE : CODEPOINT_TYPE.COMBINING_MARK);
			}
			
			for (let nIndex = 1; nIndex < nCodePointsCount - 1; ++nIndex)
			{
				this.private_HandleItem(this.Buffer[++curIndex], AscFonts.NO_GRAPHEME, _nWidth, this.FontSize, AscWord.fontslot_ASCII, isLigature ? CODEPOINT_TYPE.LIGATURE_CONTINUE : CODEPOINT_TYPE.COMBINING_MARK);
			}
		}
	};
	CParagraphTextShaper.prototype.private_CheckRun = function(oRun)
	{
		let oRunParent = oRun.GetParent();
		let oTextPr    = oRun.Get_CompiledPr(false);

		if (this.Parent !== oRunParent || !this.IsEqualTextPr(oTextPr))
			this.FlushWord();

		let oForm      = oRun.GetParentForm();
		let isCombForm = oForm && oForm.IsTextForm() && oForm.GetTextFormPr().IsComb();

		this.Paragraph = oRun.GetParagraph();
		this.Parent    = oRunParent;
		this.TextPr    = oTextPr;
		this.Spacing   = isCombForm ? 0 : oTextPr.Spacing;
		this.Ligatures = isCombForm || Math.abs(this.Spacing) > 0.001 ? Asc.LigaturesType.None : oTextPr.Ligatures;
		this.AscFont   = oRun.IsUseAscFont(oTextPr);
	};
	CParagraphTextShaper.prototype.private_HandleNBSP = function(oItem)
	{
		let oFontInfo = this.TextPr.GetFontInfo(AscWord.fontslot_ASCII);
		let nGrapheme = AscCommon.g_oTextMeasurer.GetGraphemeByUnicode(0x00B0, oFontInfo.Name, oFontInfo.Style);
		let nSpace    = AscCommon.g_oTextMeasurer.GetGraphemeByUnicode(0x0020, oFontInfo.Name, oFontInfo.Style);
		this.private_HandleItem(oItem, nGrapheme, AscFonts.GetGraphemeWidth(nSpace), oFontInfo.Size, AscWord.fontslot_ASCII, false, false, false);
	};
	CParagraphTextShaper.prototype.private_HandlePdfText = function(item)
	{
		let fontInfo = this.TextPr.GetFontInfo(AscWord.fontslot_ASCII);
		let gid = item.GetGid();
		
		let grapheme;
		if (gid)
			grapheme = AscCommon.g_oTextMeasurer.GetGraphemeByGid(gid, fontInfo.Name, fontInfo.Style, item.GetCodePoint());
		else
			grapheme = AscCommon.g_oTextMeasurer.GetGraphemeByUnicode(item.GetCodePoint(), fontInfo.Name, fontInfo.Style);
		
		let width = AscFonts.GetGraphemeWidth(grapheme);
		
		item.SetGrapheme(grapheme);
		item.SetMetrics(fontInfo.Size, AscWord.fontslot_ASCII, this.TextPr);
		item.SetWidth(width, this.TextPr, width);
		
		if (item.IsText())
			item.SetCodePointType(AscWord.CODEPOINT_TYPE.BASE);
	};
	CParagraphTextShaper.prototype.private_HandleItem = function(oItem, nGrapheme, nWidth, nFontSize, nFontSlot, nCodePointType)
	{
		if (this.Temporary)
		{
			oItem.ResetTemporaryGrapheme();
			if (nGrapheme !== oItem.GetGrapheme()
				|| nCodePointType !== oItem.GetCodePointType()
				|| Math.abs(nWidth - oItem.GetMeasuredWidth()) > 0.001)
			{
				oItem.SetTemporaryGrapheme(nGrapheme);
				oItem.SetTemporaryCodePointType(nCodePointType);
				oItem.SetTemporaryWidth(nWidth);
			}
		}
		else
		{
			oItem.SetGrapheme(nGrapheme);
			oItem.SetMetrics(nFontSize, nFontSlot, this.TextPr);
			oItem.SetCodePointType(nCodePointType);
			oItem.SetWidth(nWidth);
		}
	};
	CParagraphTextShaper.prototype.private_HandleSpace = function(item)
	{
		let codePoint = this.MaskSymbol ? this.MaskSymbol : item.GetCodePoint();
		
		let fontInfo = this.TextPr.GetFontInfo(AscWord.fontslot_ASCII);
		let grapheme = AscCommon.g_oTextMeasurer.GetGraphemeByUnicode(codePoint, fontInfo.Name, fontInfo.Style);
		
		let width   = AscFonts.GetGraphemeWidth(grapheme);
		let enWidth = getEastAsiaEnWidth(fontInfo.Name, fontInfo.Style);

		let layoutFontSize = this.private_GetLayoutFontSize(fontInfo.Size);
		if (layoutFontSize > 0)
			width = getLayoutAdvance(width, layoutFontSize);
		
		item.SetGrapheme(this.MaskSymbol ? grapheme : AscFonts.NO_GRAPHEME);
		item.SetMetrics(fontInfo.Size, AscWord.fontslot_ASCII, this.TextPr);
		item.SetWidth(width, this.TextPr, enWidth);
	};
	/**
	 * Font size (pt) at which PowerPoint's layout advances are rounded, or 0 when the
	 * advances stay linear: outside the presentation editor, and for superscript,
	 * subscript and small caps, whose size is scaled when the item is measured.
	 */
	CParagraphTextShaper.prototype.private_GetLayoutFontSize = function(fontSize)
	{
		if (!(Asc.editor && Asc.editor.isPresentationEditor) || !this.TextPr || null !== this.MaskSymbol)
			return 0;

		if (this.TextPr.VertAlign !== AscCommon.vertalign_Baseline || (this.TextPr.SmallCaps && !this.TextPr.Caps))
			return 0;

		// The size the run items store (1/64 pt), so that an item's width is a whole number of 1/8 pt.
		fontSize = ((fontSize * 64) | 0) / 64;
		return fontSize > 0 ? fontSize : 0;
	};
	/**
	 * Advance (mm per pt of font size) used to lay out one shaped glyph. PowerPoint rounds
	 * each glyph's own advance to whole 1/576-inch units (1/8 pt; a font size of S pt has 8 x S
	 * units per em) and then adds pair kerning unrounded. Measured on PowerPoint's PDF export:
	 * the start positions of words along a line follow this to within about 0.02 pt per 100 pt,
	 * while exact design widths drift by up to 0.3 pt per 100 pt.
	 */
	CParagraphTextShaper.prototype.private_GetLayoutWidth = function(grapheme, width)
	{
		let fontSize = this.private_GetLayoutFontSize(this.FontSize);
		if (fontSize <= 0)
			return width;

		let gid = AscFonts.GetGraphemeSingleGid(grapheme);
		if (gid < 0)
			return width;

		let advance = getGlyphAdvance(AscFonts.GetGraphemeFontId(grapheme), gid);
		if (advance <= 0)
			return width;

		return getLayoutAdvance(advance, fontSize) + (width - advance);
	};
	CParagraphTextShaper.prototype.IsEqualTextPr = function(oTextPr)
	{
		// Здесь мы не используем стандартную функцию CTextPr.IsEqual, потому что она сравнивает на полное
		// совпадение объектов, а нас интересует только совпадение настроек, которые не мешают сборке текста

		if (!oTextPr || !this.TextPr)
			return false;

		let t = this.TextPr;

		return (t.Bold === oTextPr.Bold
			&& t.BoldCS === oTextPr.BoldCS
			&& t.Italic === oTextPr.Italic
			&& t.ItalicCS === oTextPr.ItalicCS
			&& IsEqualNullableFloatNumbers(t.FontSize, oTextPr.FontSize)
			&& IsEqualNullableFloatNumbers(t.FontSizeCS, oTextPr.FontSizeCS)
			&& t.VertAlign === oTextPr.VertAlign
			&& IsEqualNullableFloatNumbers(t.Spacing, oTextPr.Spacing)
			&& t.SmallCaps === oTextPr.SmallCaps
			&& t.Position === oTextPr.Position
			&& t.CS === oTextPr.CS
			&& t.RTL === oTextPr.RTL
			&& t.Vanish === oTextPr.Vanish
			&& t.Ligatures === oTextPr.Ligatures
			&& t.IsKerned() === oTextPr.IsKerned()
			&& t.RFonts.IsEqualSlot(oTextPr.RFonts, this.FontSlot));
	};
	CParagraphTextShaper.prototype.GetTextScript = function(nUnicode)
	{
		// TODO: Remove it after implementing bigi algorithm
		// Check bugs 66317, 66435
		if (0x060C <= nUnicode && nUnicode <= 0x074A)
			return AscFonts.HB_SCRIPT.HB_SCRIPT_ARABIC;
		
		let script = AscFonts.hb_get_script_by_unicode(nUnicode);
		if (AscFonts.HB_SCRIPT.HB_SCRIPT_COMMON === script && this.TextPr && this.TextPr.CS)
			return AscFonts.HB_SCRIPT.HB_SCRIPT_INHERITED;
		
		return script;
	};
	CParagraphTextShaper.prototype.ShapeRunTextItem = function(item, textPr)
	{
		let fontSlot = item.GetFontSlot(textPr);
		let fontInfo = textPr.GetFontInfo(fontSlot);
		let grapheme = AscCommon.g_oTextMeasurer.GetGraphemeByUnicode(item.GetCodePoint(), fontInfo.Name, fontInfo.Style);
		item.SetGrapheme(grapheme);
		item.SetMetrics(fontInfo.Size, fontSlot, textPr);
		item.SetCodePointType(CODEPOINT_TYPE.BASE);
		item.SetWidth(AscFonts.GetGraphemeWidth(grapheme));
	};
	CParagraphTextShaper.prototype.private_IsReplaceToHindiDigits = function()
	{
		if (this.MaskSymbol)
			return false;

		if (Asc.editor.isPdfEditor())
		{
			let oParent = this.Paragraph.GetParent();
			if (oParent.ParentPDF && oParent.ParentPDF.IsForm())
				return oParent.ParentPDF.IsHindiDigits();

			return false;
		}

		let logicDocument = this.Paragraph ? this.Paragraph.GetLogicDocument() : undefined;
		return (logicDocument
			&& logicDocument.IsDocumentEditor()
			&& Asc.c_oNumeralType.hindi === logicDocument.GetNumeralType());
	};
	CParagraphTextShaper.prototype.private_HandleHindiDigit = function(oItem)
	{
		let oFontInfo = this.TextPr.GetFontInfo(AscWord.fontslot_ASCII);
		let nGrapheme = AscCommon.g_oTextMeasurer.GetGraphemeByUnicode(oItem.GetCodePoint() + (0x0660 - 0x0030), oFontInfo.Name, oFontInfo.Style);
		this.private_HandleItem(oItem, nGrapheme, AscFonts.GetGraphemeWidth(nGrapheme), oFontInfo.Size, AscWord.fontslot_ASCII, false, false, false);
	};
	CParagraphTextShaper.prototype.SetMaskSymbol = function(maskSymbol)
	{
		if (typeof maskSymbol === 'string')
			this.MaskSymbol = maskSymbol.codePointAt(0);
		else if (!isNaN(maskSymbol))
			this.MaskSymbol = maskSymbol;
		else
			this.MaskSymbol = null;
	};
	CParagraphTextShaper.prototype.GetMaskSymbol = function()
	{
		if (typeof maskSymbol === 'string')
			this.MaskSymbol = maskSymbol.codePointAt(0);
		else if (!isNaN(maskSymbol))
			this.MaskSymbol = maskSymbol;
		else
			this.MaskSymbol = null;
	};
	
	// Font key (FontNameMap id << 8 | style, as graphemes store it) -> gid -> unkerned advance
	// (mm per pt of font size), 0 when it is not known.
	let glyphAdvances = {};
	// Font key -> code point -> true once its glyph's advance has been measured.
	let prefetchedCodePoints = {};
	/**
	 * Unkerned advance of a glyph, from the cache filled by prefetchGlyphAdvances, or 0.
	 * This runs while the shaper's output is being read, so it must not call into the font
	 * engine: the engine's calls share the buffer that output is read from.
	 */
	function getGlyphAdvance(fontKey, gid)
	{
		let advances = glyphAdvances[fontKey];
		if (!advances || undefined === advances[gid])
			return 0;

		return advances[gid];
	}
	/**
	 * Measures and caches the unkerned advances of the glyphs the given code points map to,
	 * in the font the shaper is about to use, at AscFonts.MEASURE_FONTSIZE (the size the shaped
	 * advances are given at). Glyphs the shaper substitutes (ligatures, alternates) are not
	 * covered and keep their linear advance. Fonts emboldened in code are left out: their
	 * measured advance includes the emboldening, the shaped advance does not.
	 */
	function prefetchGlyphAdvances(familyName, fontStyle, codePoints)
	{
		let fontKey = AscCommon.FontNameMap.GetId(familyName) << 8 | fontStyle;
		let known   = prefetchedCodePoints[fontKey];
		if (!known)
			known = prefetchedCodePoints[fontKey] = {};

		let missing = [];
		for (let index = 0, count = codePoints.length; index < count; ++index)
		{
			if (!known[codePoints[index]])
				missing.push(codePoints[index]);
		}

		if (!missing.length)
			return;

		let advances = glyphAdvances[fontKey];
		if (!advances)
			advances = glyphAdvances[fontKey] = {};

		let measurer = AscCommon.g_oTextMeasurer;
		measurer.SetFontInternal(familyName, AscFonts.MEASURE_FONTSIZE, fontStyle);

		let font = measurer.m_oManager ? measurer.m_oManager.m_pFont : null;
		if (!font)
			return;

		let isStringGid = font.GetStringGID();
		font.SetStringGID(true);
		for (let index = 0, count = missing.length; index < count; ++index)
		{
			known[missing[index]] = true;

			let gid = font.GetGIDByUnicode(missing[index]);
			if (!gid || undefined !== advances[gid])
				continue;

			let glyph = font.m_bNeedDoBold ? null : font.GetChar(gid);
			advances[gid] = glyph && glyph.fAdvanceX > 0 ? glyph.fAdvanceX * 64 * AscFonts.GRAPHEME_COEF : 0;
		}
		font.SetStringGID(isStringGid);
	}
	/**
	 * An advance (mm per pt of font size) rounded the way PowerPoint lays text out: to a
	 * whole number of 1/8 pt at the given font size, halves rounded up.
	 */
	function getLayoutAdvance(advance, fontSize)
	{
		let eighths = Math.floor(advance * fontSize * 72 / 25.4 * 8 + 0.5 + 1e-6);
		return eighths / 8 * 25.4 / 72 / fontSize;
	}

	let enWidth = {};
	function getEastAsiaEnWidth(fontName, fontStyle)
	{
		if (!enWidth[fontName])
			enWidth[fontName] = {};
		
		if (undefined !== enWidth[fontName][fontStyle])
			return enWidth[fontName][fontStyle];
		
		function getWidth(codePoint, flags)
		{
			let grapheme = AscCommon.g_oTextMeasurer.GetGraphemeByUnicode(codePoint, fontName, fontStyle);
			if (AscFonts.NO_GRAPHEME === grapheme)
				return 0;
			
			let fontId = AscFonts.GetGraphemeFontId(grapheme);
			let fName  = AscFonts.GetFontNameByFontId(fontId);
			let fStyle = AscFonts.GetFontStyleByFontId(fontId);
			
			let info = AscCommon.g_oTextMeasurer.GetFontBySymbol(codePoint, null, false);
			let fontFace = info.Font && info.Font && info.Font.m_pFaceInfo ? info.Font.m_pFaceInfo : null;
			if (info.CodePoint !== codePoint
				|| !fontFace
				|| fName !== fontFace.family_name
				|| fStyle !== fontStyle
				|| 0 === (fontFace.os2_ulCodePageRange1 & flags))
				return 0;
			
			return AscFonts.GetGraphemeWidth(grapheme) / 2;
		}
		
		let w = getWidth(0xE400, 0x160000); // (1 << 17) | (1 << 18) | (1 << 20);
		if (0 !== w)
			return w;
		
		w = getWidth(0xAC00, 0x280000); // (1 << 19) | (1 << 21);
		w = 0 !== w ? w : 25.4 / 72 / 2;
		enWidth[fontName][fontStyle] = w;
		return w;
	}
	
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CODEPOINT_TYPE      = CODEPOINT_TYPE;
	window['AscWord'].ParagraphTextShaper = new CParagraphTextShaper();
	window['AscWord'].stringShaper        = new AscFonts.StringShaper();

})(window);
