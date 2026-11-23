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
	/**
	 * Class for storing the current draw state of paragraph content (text, images)
	 * @param {AscWord.ParagraphDrawState} drawState - reference to the main state
	 * @constructor
	 */
	function ParagraphContentDrawState(drawState)
	{
		this.drawState = drawState;
		
		this.Paragraph = undefined;
		this.Graphics  = undefined;
		this.BgColor   = undefined;
		
		this.Theme     = undefined;
		this.ColorMap  = undefined;
		
		this.CurPos = new AscWord.CParagraphContentPos();
		
		this.VisitedHyperlink = false;
		this.Hyperlink = false;
		
		this.Page   = 0;
		this.Line   = 0;
		this.Range  = 0;
		
		this.X = 0;
		this.Y = 0;
		
		this.LineTop    = 0;
		this.LineBottom = 0;
		this.BaseLine   = 0;
		
		this.complexFields = new AscWord.ParagraphComplexFieldStack();
		
		this.run            = null;
		this.yOffset        = 0;
		this.textPr         = null;
		this.reviewColor    = null;
		this.themeColor     = null;
		this.autoColor      = null;
		this.calcY          = 0; // calculated vertical position taking into account sub/super script
		
		// for Math
		this.mathTextInfo = null;
		this.paraMath     = null;
		
		// Slide text: where PowerPoint draws the glyphs across the slide (see getSlideGlyphX)
		this.slideRuns     = null;  // {transform, stretch, offset, grid, extras} for this line, or null
		this.slideRunStart = true;  // the next glyph starts a glyph run
		this.slideShift    = 0;     // how far from its layout position the glyph run's next glyph is drawn, in mm
		this.slidePhase    = 0;     // the sixteenth of a grid step its start was rounded off, -8..7
		this.slideVisible  = false; // whether it has had a glyph other than a space
		this.slideExtras   = false; // whether it gets the extra steps
		this.slideAfterDash = false; // the next glyph run follows a hyphen or a dash
		this.slideLineGlyph = false; // the line has had a glyph other than a space
		this.slideLastSpace = false; // the line's last glyph was a space
		this.slideDevice    = false; // the glyph run is drawn from its device positions (getSlideDeviceX)
		this.slideRunX      = 0;     // the layout position of its start, in mm
		this.slideCount     = 0;     // how many glyphs it has had
		this.slidePrev      = null;  // its last glyph: {dev, nat, kern, isSpace, w, u} (getSlideDeviceX)
		this.slidePdf       = 0;     // where its PDF draws the glyph, in grid steps from its start
		this.slideAcc       = 0;     // what its PDF writer has not yet written out, in 1/1000 em of the PDF's font size
		this.slideLineY     = null;  // the line's rounded baseline down the slide before the stretch, in mm, or null

		this.bidiFlow = new AscWord.BidiFlow(this);
	}
	ParagraphContentDrawState.prototype.init = function()
	{
		this.Paragraph = this.drawState.getParagraph();
		this.Graphics  = this.drawState.getGraphics();
		
		this.BgColor  = this.drawState.getBgColor();
		this.Theme    = this.drawState.getTheme();
		this.ColorMap = this.drawState.getColorMap();
		
		this.VisitedHyperlink = false;
		this.Hyperlink        = false;
		
		this.CurPos = new AscWord.CParagraphContentPos();
	};
	ParagraphContentDrawState.prototype.resetPage = function(page)
	{
		this.Page = page;
		this.complexFields.resetPage(this.Paragraph, page);
	};
	ParagraphContentDrawState.prototype.resetLine = function(line, baseLine, lineTop, lineBottom)
	{
		this.Line = line;
		
		this.LineTop    = lineTop;
		this.LineBottom = lineBottom;
		this.BaseLine   = baseLine;
		
		this.Y = baseLine;
	};
	/**
	 * @param info {?{transform: AscCommon.CMatrix, stretch: number, offset: number, grid: number, extras: boolean,
	 *        stretchY: number}} - Paragraph.GetSlideTextGridX's grid for this line, null for text drawn at its layout
	 *        positions
	 * @param [lineY=null] {?number} - the line's rounded baseline down the slide before the stretch, in mm
	 *        (Paragraph.GetSlideLinePos), null when it is not known
	 */
	ParagraphContentDrawState.prototype.setSlideGlyphRuns = function(info, lineY)
	{
		this.slideRuns      = info;
		this.slideLineY     = (undefined !== lineY ? lineY : null);
		this.slideRunStart  = true;
		this.slideAfterDash = false;
		this.slideLineGlyph = false;
		this.slideLastSpace = false;
	};
	/**
	 * @param range {number}
	 * @param x {number}
	 * @param [numShift=0] {number} - slide text: how much further on than its layout the text after the numbering starts
	 */
	ParagraphContentDrawState.prototype.beginRange = function(range, x, numShift)
	{
		this.run = null;
		
		this.Range = range;
		
		this.X = x;
		
		this.slideRunStart  = true;
		this.slideAfterDash = false;
		this.slideLineGlyph = false;
		this.slideLastSpace = false;
		
		let isRtl = this.Paragraph.isRtlDirection();
		if (this.Paragraph.Numbering.checkRange(this.Range, this.Line) && !isRtl)
		{
			this.handleNumbering();
			if (numShift)
				this.X += numShift;
		}
		
		this.bidiFlow.begin(this.Paragraph.isRtlDirection());
	};
	ParagraphContentDrawState.prototype.endRange = function()
	{
		this.bidiFlow.end();
		
		let isRtl = this.Paragraph.isRtlDirection();
		if (this.Paragraph.Numbering.checkRange(this.Range, this.Line) && isRtl)
			this.handleNumbering();
	};
	/**
	 * @param element {AscWord.CRunElementBase}
	 * @param run {AscWord.CRun}
	 */
	ParagraphContentDrawState.prototype.handleRunElement = function(element, run)
	{
		if (!this.complexFields.checkRunElement(element))
			return;
		
		this.bidiFlow.add([element, run], element.getBidiType());
	};
	ParagraphContentDrawState.prototype.handleBidiFlow = function(data, direction)
	{
		let element = data[0];
		this.handleRun(data[1]);
		
		switch (element.Type)
		{
			case para_Text:
				this.handleText(element, direction);
				break;
			case para_Drawing:
				this.handleDrawing(element);
				break;
			case para_End:
				this.handleParagraphMark(element);
				break;
			case para_Math_Ampersand:
			case para_Math_Text:
			case para_Math_BreakOperator:
			case para_Math_Placeholder:
				this.handleMathElement(element);
				break;
			case para_FieldChar:
				this.handleFieldChar(element);
				break;
			default:
				this.handleRegularElement(element, direction);
				break;
		}
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	/**
	 * @param run {AscWord.CRun}
	 */
	ParagraphContentDrawState.prototype.handleRun = function(run)
	{
		if (this.run === run)
			return;
		
		this.run = run;
		
		// PowerPoint draws each a:r as a glyph run of its own, but one of only spaces goes on into the next
		if (this.slideVisible)
			this.slideRunStart = true;
		
		this.yOffset = run.getYOffset();
		
		let textPr = run.getCompiledPr();
		
		if (run.IsMathRun())
		{
			this.mathTextInfo = new CMathInfoTextPr({
				TextPr      : textPr,
				ArgSize     : run.Parent.Compiled_ArgSz.value,
				bNormalText : run.IsNormalText(),
				bEqArray    : run.bEqArray
			});
			this.paraMath = run.ParaMath;
		}
		else
		{
			this.mathTextInfo = null;
			this.paraMath     = null;
		}
		
		this.reviewColor = null;
		this.themeColor  = null;
		this.autoColor   = null;
		
		if (reviewtype_Common !== run.GetReviewType())
		{
			this.reviewColor = run.GetReviewColor();
		}
		else if (this.VisitedHyperlink)
		{
			AscFormat.G_O_VISITED_HLINK_COLOR.check(this.Theme, this.ColorMap);
			let RGBA = AscFormat.G_O_VISITED_HLINK_COLOR.getRGBAColor();
			this.themeColor = new CDocumentColor(RGBA.R, RGBA.G, RGBA.B, false);
		}
		else if (textPr.Unifill)
		{
			let RGBA = null;
			if (this.isSlideEditor() && this.Hyperlink)
			{
				AscFormat.G_O_HLINK_COLOR.check(this.Theme, this.ColorMap);
				RGBA = AscFormat.G_O_HLINK_COLOR.getRGBAColor();
			}
			else if (!this.isTextArtDraw())
			{
				textPr.Unifill.check(this.Theme, this.ColorMap);
				RGBA = textPr.Unifill.getRGBAColor();
			}
			
			if (RGBA)
				this.themeColor = new CDocumentColor(RGBA.R, RGBA.G, RGBA.B, false);
		}
		
		if (textPr.FontRef && textPr.FontRef.Color)
		{
			textPr.FontRef.Color.check(this.Theme, this.ColorMap);
			let RGBA = textPr.FontRef.Color.RGBA;
			this.autoColor = new CDocumentColor(RGBA.R, RGBA.G, RGBA.B, false);
		}
		else
		{
			let bgColor = this.BgColor;
			if (textPr.Shd && !textPr.Shd.IsNil())
				bgColor = textPr.Shd.GetSimpleColor(this.Theme, this.ColorMap);
			
			this.autoColor = (bgColor && !bgColor.Check_BlackAutoColor() ? new CDocumentColor(255, 255, 255, false) : new CDocumentColor(0, 0, 0, false));
		}
		
		this.updateGraphicsState(textPr, run.IsUseAscFont(textPr));
		
		this.calcY = this.calculateY(textPr);
	};
	/**
	 * @param textPr {AscWord.CTextPr}
	 * @param [useAscFont=false] {boolean}
	 */
	ParagraphContentDrawState.prototype.updateGraphicsState = function(textPr, useAscFont)
	{
		if (textPr)
		{
			this.textPr = textPr;
			
			if (useAscFont)
			{
				textPr = textPr.Copy();
				textPr.RFonts.SetAll("ASCW3", -1);
			}
			
			this.Graphics.SetTextPr(textPr, this.Theme);
		}
		
		let color = this.getTextColor();
		if (color)
		{
			this.Graphics.b_color1(color.r, color.g, color.b, 255);
			this.Graphics.p_color(color.r, color.g, color.b, 255);
		}
	};
	ParagraphContentDrawState.prototype.getTextColor = function()
	{
		if (this.reviewColor)
			return this.reviewColor;
		if (this.themeColor)
			return this.themeColor;
		
		if (this.isTextArtDraw())
		{
			if (this.textPr.Color.IsAuto() && !this.textPr.TextFill)
				return this.autoColor;
			return null;
		}
		
		if (this.textPr.Color.IsAuto())
			return this.autoColor;
		
		return this.textPr.Color;
	};
	/**
	 * @param text {AscWord.CRunText}
	 * @param direction {AscBidi.DIRECTION}
	 */
	ParagraphContentDrawState.prototype.handleText = function(text, direction)
	{
		let x = this.X;
		if (this.slideRuns)
		{
			let codePoint = text.GetCodePoint();
			
			// PowerPoint draws a hyphen, an en dash or an em dash as a glyph run of its own
			let isDash = (0x002D === codePoint || 0x2013 === codePoint || 0x2014 === codePoint);
			if (isDash || AscBidi.DIRECTION.R === direction)
				this.slideRunStart = true;
			
			if (AscBidi.DIRECTION.R !== direction)
				x = this.getSlideGlyphX(text.GetWidthVisible(), 0x00A0 !== codePoint, AscWord.GetSlideGlyphAdvance(text.GetGrapheme()), text.GetFontSize());
			
			if (isDash)
			{
				this.slideRunStart  = true;
				this.slideAfterDash = true;
			}
		}
		
		if (AscBidi.DIRECTION.R === direction && AscBidi.isPairedBracket(text.GetCodePoint()))
			text.Draw(x, this.calcY - this.yOffset, this.Graphics, this, this.textPr, AscBidi.getPairedBracketGrapheme(text.Grapheme));
		else
			text.Draw(x, this.calcY - this.yOffset, this.Graphics, this, this.textPr);
		
		this.X += text.GetWidthVisible();
	};
	/**
	 * Slide text (I-027): where PowerPoint's PDF draws a glyph across the slide, for its layout position this.X.
	 * Its PDF puts text positions on a 1/3000-inch grid, stretched about the slide's left edge (GetSlideTextGridX). It
	 * starts each glyph run (each a:r, and each hyphen, en dash or em dash on its own) at its layout position stretched
	 * and rounded to that grid, and keeps the sixteenths of a step the start was rounded down by (-8..7). Each glyph
	 * then advances by its own unstretched width, plus a step whenever its stretch (in sixteenths of a step, rounded)
	 * and those sixteenths reach half a step. As the sixteenths are not carried on from glyph to glyph, a run whose
	 * start was rounded down by 7/16 has every glyph wider than about 5 pt drawn a step further on than the one
	 * before, which makes the run up to 2 per mille longer than its layout (at 6/16, glyphs wider than about 15 pt);
	 * most runs are drawn at their layout widths. Spaces before the first other glyph of a run do carry the sixteenths
	 * on at the start of a line (an a:r of only spaces going on into the next a:r); later in a line they are drawn on
	 * their own. Only left and justified lines (extras), whose start is known to a sixteenth, get the extra steps, and
	 * of those only the line's first glyph run and glyph runs that follow another glyph straight on (an a:r boundary
	 * inside a word or before punctuation): not those after a space, a hyphen or a dash, whose sixteenths are not
	 * known, nor runs with character spacing, which PowerPoint draws with the PDF's Tc.
	 * Inside a glyph run the glyphs go where PowerPoint's device puts them and its PDF writer writes them
	 * (getSlideDeviceX, I-024), which draws them a little closer together than the layout; runs with character
	 * spacing, and glyphs whose own advance is not known, are drawn at their layout widths.
	 * @param width {number} - the glyph's width in the layout, in mm
	 * @param isVisible {boolean} - whether it is a glyph other than a space
	 * @param [advance=0] {number} - its font's unkerned advance, in mm per pt of font size (0 when not known)
	 * @param [fontSize=0] {number} - the font size it is drawn at, in pt
	 * @returns {number} the x to draw the glyph at
	 */
	ParagraphContentDrawState.prototype.getSlideGlyphX = function(width, isVisible, advance, fontSize)
	{
		let info = this.slideRuns;
		let t    = info.transform;
		let grid = info.grid;

		let hasAdvance = (advance > 0 && fontSize > 0);

		// spaces that open a glyph run after the line's first glyph are drawn on their own: the glyph after them
		// starts the glyph run again, without the extra steps (PowerPoint drew 1 of 36 such runs long)
		let isRestart = (!this.slideRunStart && isVisible && !this.slideVisible && this.slideLineGlyph);
		// a glyph whose advance is not known in a run drawn from its device positions starts a glyph run of its own
		let isUnknown = (!this.slideRunStart && this.slideDevice && !hasAdvance);
		if (this.slideRunStart || isRestart || isUnknown)
		{
			this.slideRunStart = false;

			// the start in sixteenths of a step, on the slide
			let slideX = t.TransformPointX(this.X, this.Y);
			let pos16  = Math.floor(16 * (slideX * (1 + info.stretch) / grid + info.offset));
			let start  = Math.floor((pos16 + 8) / 16);

			this.slidePhase   = pos16 - 16 * start;
			this.slideShift   = (start * grid - slideX) / t.sx;
			this.slideVisible = false;

			// later in a line, a glyph run after a space, a hyphen or a dash does not get them either (on the 1646-slide
			// deck PowerPoint drew 6 of 80 such runs long, and 523 of 572 after another glyph)
			let afterSpace = this.slideLineGlyph && this.slideLastSpace;
			this.slideExtras    = info.extras && !this.slideAfterDash && !afterSpace && !isRestart && !isUnknown;
			this.slideAfterDash = false;

			this.slideDevice = hasAdvance && !this.textPr.Spacing;
			this.slideRunX   = this.X;
			this.slideCount  = 0;
			this.slidePrev   = null;
			this.slidePdf    = 0;
			this.slideAcc    = 0;
		}

		let x = this.slideDevice ? this.getSlideDeviceX(width, isVisible, advance, fontSize) : this.X + this.slideShift;

		if (isVisible)
		{
			this.slideVisible   = true;
			this.slideLineGlyph = true;
		}
		this.slideLastSpace = !isVisible;

		// runs with character spacing (spc) are drawn with the PDF's Tc and do not get the extra steps
		if (!this.slideDevice && this.slideExtras && !this.textPr.Spacing)
		{
			let stretch16 = Math.floor(16 * width * t.sx / grid * info.stretch + 0.5);
			let steps     = Math.floor((this.slidePhase + 8 + stretch16) / 16);
			if (steps > 0)
				this.slideShift += steps * grid / t.sx;
			
			if (!this.slideVisible)
				this.slidePhase += stretch16 - 16 * steps;
		}

		return x;
	};
	/**
	 * Slide text (I-024): where PowerPoint's PDF draws a glyph of a glyph run (getSlideGlyphX), from where its device
	 * puts it and how its PDF writer writes that. Measured on PowerPoint's PDF export of the 1646-slide deck (Calibri
	 * 12.75, 13.5 and 18 pt) and of a test deck (Calibri 12 to 32 pt, glyphs repeated and a sentence, in three boxes):
	 * - The device works in whole pixels of 600 dpi, at a font size of round(size * 600 / 72) pixels per em. A glyph
	 *   after a space (a word start) goes to the pixel nearest its layout position from the run's start (an exact half
	 *   goes down below 1024 pixels and up from there). Any other glyph follows the one before by that one's advance in
	 *   whole pixels plus its kerning in whole pixels, and is moved back a pixel at a time while it is more than 3.5
	 *   pixels ahead of its layout position, and on a pixel at a time while its end (its start and its own advance in
	 *   pixels) is 5.5 pixels or more short of the next glyph's layout position.
	 * - The PDF writer draws each glyph by its width in the PDF's font (1/1000 em: the advance rounded to 1/64 and then
	 *   to whole units) at a font size of 5 grid steps a pixel per em, stretched down the slide from the baseline (a
	 *   step larger in some lines). It adds up how far the device's advances (with their extra steps) are ahead
	 *   of those widths, and when that reaches 3 units either way it moves the next glyph by it, less a half, cut
	 *   towards zero, and starts again from nothing: what it leaves out is lost, so glyphs end up slightly closer
	 *   together than the device has them (about 0.27 per mille at 13.5 pt, 0.5 at 18 pt, 1.3 at 12.75 pt).
	 * Spaces before the first other glyph of the run are not written: the glyph after them starts the PDF's text there.
	 * Glyph runs are measured from their own start, not the line's. On the deck this gives PowerPoint's moves exactly
	 * in 63% of lines at 13.5 pt, 32% at 18 pt and 95% at 12.75 pt, and the line ends to within 0.06 per mille; the
	 * rest are word starts a pixel off, which PowerPoint rounds a little unevenly.
	 * @param width {number} - the glyph's width in the layout, in mm
	 * @param isVisible {boolean} - whether it is a glyph other than a space
	 * @param advance {number} - its font's unkerned advance, in mm per pt of font size
	 * @param fontSize {number} - the font size it is drawn at, in pt
	 * @returns {number} the x to draw the glyph at
	 */
	ParagraphContentDrawState.prototype.getSlideDeviceX = function(width, isVisible, advance, fontSize)
	{
		let info  = this.slideRuns;
		let t     = info.transform;
		let grid  = info.grid;
		let steps = t.sx / grid; // grid steps (1/3000 inch, 5 a pixel) per mm of the layout

		// its layout position from the run's start (item widths are kept cut to 1/16384 mm: half of that back for each)
		let ref = (this.X - this.slideRunX + this.slideCount * 0.5 / AscWord.TEXTWIDTH_DIVIDER) * steps;

		let ppem = Math.floor(fontSize * 600 / 72 + 0.5);
		let em   = advance * 72 / 25.4;
		let nat  = 5 * Math.floor(em * ppem + 0.5 + 1e-9);

		let prev = this.slidePrev;
		let dev  = 0;
		if (prev && isVisible && prev.isSpace)
		{
			let px   = ref / 5;
			let half = Math.abs(px - Math.floor(px) - 0.5) < 1e-3;
			dev = 5 * (half ? (px < 1024 ? Math.floor(px) : Math.ceil(px)) : Math.floor(px + 0.5));
		}
		else if (prev)
		{
			dev = prev.dev + prev.nat + prev.kern;
			while (dev - ref > 17.5)
				dev -= 5;

			let nextRef = ref + width * steps;
			while (nextRef - (dev + nat) >= 27.5)
				dev += 5;
		}

		if (prev)
		{
			let devAdvance = dev - prev.dev;
			let pdfAdvance = devAdvance;
			if (this.slideExtras)
			{
				let stretch16 = Math.floor(16 * devAdvance * info.stretch + 0.5);
				let extra     = Math.floor((this.slidePhase + 8 + stretch16) / 16);
				pdfAdvance += extra;
				if (!this.slideVisible)
					this.slidePhase += stretch16 - 16 * extra;
			}

			if (!this.slideVisible)
			{
				this.slidePdf += pdfAdvance;
			}
			else
			{
				let acc  = Math.round((this.slideAcc + pdfAdvance / prev.u - prev.w) * 1e9) / 1e9;
				let move = 0;
				if (Math.abs(acc) >= 3)
				{
					move = Math.trunc(acc - 0.5);
					acc  = 0;
				}
				this.slideAcc  = acc;
				this.slidePdf += (prev.w + move) * prev.u;
			}
		}

		// the PDF's font size in grid steps: the font size in pixels stretched down the slide from the baseline, as the
		// baseline is (GetSlideLinePos), so a step larger in some lines (Calibri 13.5 pt: 13.584 for 13.56 in one line
		// in eight); its width of the glyph (1/1000 em) and the kerning after it in the layout
		let tf = 5 * ppem;
		if (null !== this.slideLineY && info.stretchY)
		{
			let lineY = this.slideLineY / grid;
			tf = Math.floor((lineY + tf) * (1 + info.stretchY) + 0.5) - Math.floor(lineY * (1 + info.stretchY) + 0.5);
		}
		let w    = Math.floor((Math.floor(em * 64000 + 0.5 + 1e-6) + 32) / 64);
		let kern = (width * 72 / 25.4 - Math.floor(em * 8 * fontSize + 0.5 + 1e-6) / 8) / fontSize;

		this.slidePrev = {
			dev     : dev,
			nat     : nat,
			kern    : 5 * Math.floor(kern * ppem + 0.5),
			isSpace : !isVisible,
			w       : w,
			u       : tf / 1000
		};
		this.slideCount++;

		return this.slideRunX + this.slideShift + this.slidePdf * grid / t.sx;
	};
	/**
	 * @param drawing {ParaDrawing}
	 */
	ParagraphContentDrawState.prototype.handleDrawing = function(drawing)
	{
		if (!drawing.IsInline())
			return;
		
		this.slideRunStart = true;
		
		drawing.Draw(this.X, this.calcY - this.yOffset, this.Graphics, this, this.textPr);
		this.X += drawing.GetWidthVisible();
		
		// Внутри отрисовки инлайн-автофигур могут изменится цвета и шрифт, поэтому восстанавливаем настройки
		this.updateGraphicsState();
	};
	/**
	 * @param element {AscWord.CRunElementBase}
	 * @param [direction] {AscBidi.DIRECTION}
	 */
	ParagraphContentDrawState.prototype.handleRegularElement = function(element, direction)
	{
		let x = this.X;
		if (this.slideRuns)
		{
			// a space is a glyph of the glyph run; a tab or anything else ends it
			if (para_Space === element.Type && AscBidi.DIRECTION.R !== direction)
			{
				let fontInfo = this.textPr.GetFontInfo(AscWord.fontslot_ASCII);
				x = this.getSlideGlyphX(element.GetWidthVisible(), false, AscWord.GetSlideSpaceAdvance(fontInfo.Name, fontInfo.Style), element.GetFontSize());
			}
			else
				this.slideRunStart = true;
		}
		
		element.Draw(x, this.calcY - this.yOffset, this.Graphics, this, this.textPr);
		this.X += element.GetWidthVisible();
		
		if (element.IsTab())
			this.updateGraphicsState();
	};
	/**
	 * @param paraMark {AscWord.CRunParagraphMark}
	 */
	ParagraphContentDrawState.prototype.handleParagraphMark = function(paraMark)
	{
		if (this.Paragraph.IsInFixedForm())
			return;
		
		let sectPr = this.Paragraph.Get_SectionPr();
		let logicDocument = this.Paragraph.GetLogicDocument();
		if (!logicDocument)
		{
			this.X += paraMark.GetWidthVisible();
			return;
		}
		
		if (logicDocument !== this.Paragraph.GetParent())
			sectPr = undefined;
		
		let editor = logicDocument.GetApi();
		if ((!editor || !editor.ShowParaMarks) && (sectPr || !this.reviewColor))
		{
			this.X += paraMark.GetWidthVisible();
			return;
		}
		
		let y = this.Y;
		if (!sectPr)
		{
			let endTextPr = true !== this.Graphics.m_bIsTextDrawer ? this.Paragraph.GetParaEndCompiledPr() : this.textPr;
			this.updateGraphicsState(endTextPr);
			y = this.calculateY(endTextPr);
		}
		
		paraMark.Draw(this.X, y - this.yOffset, this.Graphics);
		this.X += paraMark.GetWidthVisible();
	};
	/**
	 * @param element {CMathText | CMathAmp}
	 */
	ParagraphContentDrawState.prototype.handleMathElement = function(element)
	{
		if (para_Math_Placeholder === element.Type && this.Graphics.isPdf())
			return;
		
		this.slideRunStart = true;
		
		let linePos = this.paraMath.GetLinePosition(this.Line, this.Range);
		element.Draw(linePos.x, linePos.y, this.Graphics, this.mathTextInfo);
		this.X += element.GetWidthVisible();
	};
	/**
	 * @param fieldChar {ParaFieldChar}
	 */
	ParagraphContentDrawState.prototype.handleFieldChar = function(fieldChar)
	{
		if (!fieldChar.IsVisual())
			return;
		
		// Draw the auto-calculated pageNum in the header/footer
		let complexField = fieldChar.GetComplexField();
		
		// We can't use textPr of the current run, since the properties of the actual text
		// may be different
		let textPr = complexField.GetFieldValueTextPr(true);
		this.updateGraphicsState(textPr);
		
		this.handleRegularElement(fieldChar);
	};
	ParagraphContentDrawState.prototype.handleNumbering = function()
	{
		let graphics = this.Graphics;
		let BgColor  = this.drawState.getBgColor();
		let Pr       = this.drawState.getParagraphCompiledPr();
		let isRtl    = Pr.ParaPr.Bidi;
		
		let X = this.X;
		let Y = this.Y;
		let CurLine = this.Line;
		let CurPage = this.Page;
		
		let para = this.Paragraph;
		
		var nReviewType  = para.GetReviewType();
		var oReviewColor = para.GetReviewColor();
		
		let numItem = para.Numbering;
		if (para_Numbering === numItem.Type)
		{
			var isHavePrChange = para.HavePrChange();
			
			let prevNumPr = para.GetPrChangeNumPr();
			let numPr     = Pr.ParaPr.NumPr;
			
			if (prevNumPr && prevNumPr.IsValid())
			{
				prevNumPr = prevNumPr.Copy();
				if (undefined === prevNumPr.Lvl)
					prevNumPr.Lvl = 0;
			}
			else
			{
				prevNumPr = null;
			}
			
			var isHaveNumbering = false;
			if ((undefined === para.Get_SectionPr() || true !== para.IsEmpty())
				&& (!para.Parent || !para.Parent.IsEmptyParagraphAfterTableInTableCell(para.GetIndex()))
				&& ((numPr && numPr.IsValid()) || (prevNumPr && prevNumPr.IsValid())))
			{
				isHaveNumbering = true;
			}
			
			if (!isHaveNumbering || (!numPr && !prevNumPr))
			{
				// Ничего не делаем
			}
			else
			{
				var oNumbering = para.Parent.GetNumbering();
				
				let oNumLvl  = null;
				let nNumSuff = Asc.c_oAscNumberingSuff.None;
				if (numPr)
				{
					oNumLvl  = oNumbering.GetNum(numPr.NumId).GetLvl(numPr.Lvl);
					nNumSuff = oNumLvl.GetSuff();
				}
				else if (prevNumPr)
				{
					// MSWord uses tab instead of suff from PrevNum (74525)
					oNumLvl  = oNumbering.GetNum(prevNumPr.NumId).GetLvl(prevNumPr.Lvl);
					nNumSuff = Asc.c_oAscNumberingSuff.Tab;
				}
				
				var nNumJc     = oNumLvl.GetJc();
				var oNumTextPr = para.GetNumberingTextPr();
				
				var oPrevNumTextPr = prevNumPr ? para.Get_CompiledPr2(false).TextPr.Copy() : null;
				if (oPrevNumTextPr && (prevNumPr && prevNumPr.IsValid()))
				{
					var oPrevNumLvl = oNumbering.GetNum(prevNumPr.NumId).GetLvl(prevNumPr.Lvl);
					oPrevNumTextPr.Merge(para.TextPr.Value.Copy());
					oPrevNumTextPr.Merge(oPrevNumLvl.GetTextPr());
				}
				
				var X_start = this.X;
				
				let numWidth  = numItem.getNumWidth();
				let suffWidth = numItem.getSuffWidth();
				if (isRtl)
				{
					X_start = this.X + numItem.getVisibleWidth();
					if (align_Left === nNumJc)
						X_start -= numWidth;
					else if (align_Center === nNumJc)
						X_start -= numWidth / 2;
				}
				else
				{
					if (align_Right === nNumJc)
						X_start -= numWidth;
					else if (align_Center === nNumJc)
						X_start -= numWidth / 2;
				}
				
				var AutoColor = ( undefined !== BgColor && false === BgColor.Check_BlackAutoColor() ? new CDocumentColor(255, 255, 255, false) : new CDocumentColor(0, 0, 0, false) );
				
				var RGBA;
				if (oNumTextPr.Unifill)
				{
					oNumTextPr.Unifill.check(this.Theme, this.ColorMap);
					RGBA = oNumTextPr.Unifill.getRGBAColor();
					graphics.b_color1(RGBA.R, RGBA.G, RGBA.B, 255);
				}
				else
				{
					if (true === oNumTextPr.Color.Auto)
					{
						if(oNumTextPr.FontRef && oNumTextPr.FontRef.Color)
						{
							oNumTextPr.FontRef.Color.check(this.Theme, this.ColorMap);
							RGBA = oNumTextPr.FontRef.Color.RGBA;
							graphics.b_color1(RGBA.R, RGBA.G, RGBA.B, 255);
						}
						else
						{
							graphics.b_color1(AutoColor.r, AutoColor.g, AutoColor.b, 255);
						}
					}
					else
					{
						graphics.b_color1(oNumTextPr.Color.r, oNumTextPr.Color.g, oNumTextPr.Color.b, 255);
					}
				}
				
				if (numItem.havePrevNumbering() || reviewtype_Common !== nReviewType)
				{
					if (reviewtype_Common === nReviewType)
						graphics.b_color1(REVIEW_NUMBERING_COLOR.r, REVIEW_NUMBERING_COLOR.g, REVIEW_NUMBERING_COLOR.b, 255);
					else
						graphics.b_color1(oReviewColor.r, oReviewColor.g, oReviewColor.b, 255);
				}
				else if (isHavePrChange && numPr && !prevNumPr)
				{
					var oPrReviewColor = para.GetPrReviewColor();
					graphics.b_color1(oPrReviewColor.r, oPrReviewColor.g, oPrReviewColor.b, 255);
				}
				
				let baseY = Y;
				switch (oNumTextPr.VertAlign)
				{
					case AscCommon.vertalign_SubScript:
					{
						baseY -= AscCommon.vaKSub * oNumTextPr.FontSize * g_dKoef_pt_to_mm;
						break;
					}
					case AscCommon.vertalign_SuperScript:
					{
						baseY -= AscCommon.vaKSuper * oNumTextPr.FontSize * g_dKoef_pt_to_mm;
						break;
					}
				}
				
				// Рисуется только сам символ нумерации
				switch (nNumJc)
				{
					case align_Right:
						numItem.Draw(X_start, baseY, graphics, oNumbering, oNumTextPr, this.Theme, oPrevNumTextPr, isRtl);
						break;
					
					case align_Center:
						numItem.Draw(X_start, baseY, graphics, oNumbering, oNumTextPr, this.Theme, oPrevNumTextPr, isRtl);
						break;
					
					case align_Left:
					default:
						numItem.Draw(X_start, baseY, graphics, oNumbering, oNumTextPr, this.Theme, oPrevNumTextPr, isRtl);
						break;
				}
				
				if (true === oNumTextPr.Strikeout || true === oNumTextPr.Underline)
				{
					if (oNumTextPr.Unifill)
					{
						graphics.p_color(RGBA.R, RGBA.G, RGBA.B, 255);
					}
					else
					{
						if (true === oNumTextPr.Color.Auto)
							graphics.p_color(AutoColor.r, AutoColor.g, AutoColor.b, 255);
						else
							graphics.p_color(oNumTextPr.Color.r, oNumTextPr.Color.g, oNumTextPr.Color.b, 255);
					}
				}
				
				let lineW      = (oNumTextPr.FontSize / 18) * g_dKoef_pt_to_mm;
				let strikeoutY = (baseY - oNumTextPr.FontSize * g_dKoef_pt_to_mm * 0.27);
				let underlineY = (baseY + para.Lines[CurLine].Metrics.TextDescent * 0.4);
				if (numItem.havePrevNumbering() || reviewtype_Common !== nReviewType)
				{
					let prevNumWidth = numItem.getPrevNumWidth();
					if (reviewtype_Common === nReviewType)
						graphics.p_color(REVIEW_NUMBERING_COLOR.r, REVIEW_NUMBERING_COLOR.g, REVIEW_NUMBERING_COLOR.b, 255);
					else
						graphics.p_color(oReviewColor.r, oReviewColor.g, oReviewColor.b, 255);
					
					// Либо у нас есть удаленная часть, либо у нас одновременно добавлен и удален параграф, тогда мы зачеркиваем суффикс
					if (numItem.havePrevNumbering() || (!numItem.havePrevNumbering() && !numItem.haveFinalNumbering()))
					{
						if (numItem.haveFinalNumbering() || nReviewType !== reviewtype_Remove)
						{
							let x0 = X_start;
							let x1 = X_start + prevNumWidth;
							graphics.drawHorLine(0, strikeoutY, x0, x1, lineW);
						}
						else
						{
							let x0 = isRtl ? X_start - suffWidth : X_start;
							let x1 = x0 + numWidth + suffWidth;
							graphics.drawHorLine(0, strikeoutY, x0, x1, lineW);
						}
					}
					
					if (numItem.haveFinalNumbering())
					{
						if (isRtl)
						{
							let x0 = X_start - suffWidth;
							let x1 = X_start;
							graphics.drawHorLine(0, underlineY, x0, x1, lineW);
							
							x0 = X_start + prevNumWidth;
							x1 = X_start + numWidth;
							graphics.drawHorLine(0, underlineY, x0, x1, lineW);
						}
						else
						{
							let x0 = X_start + prevNumWidth;
							let x1 = X_start + numWidth + suffWidth;
							graphics.drawHorLine(0, underlineY, x0, x1, lineW);
						}
					}
				}
				else if (isHavePrChange && numPr && !prevNumPr)
				{
					let prevColor = para.GetPrReviewColor();
					graphics.p_color(prevColor.r, prevColor.g, prevColor.b, 255);
					
					let x0 = isRtl ? X_start - suffWidth : X_start;
					let x1 = x0 + numWidth + suffWidth;
					graphics.drawHorLine(0, underlineY, x0, x1, lineW);
				}
				else
				{
					let x0 = X_start;
					let x1 = x0 + numWidth;
					
					if (oNumTextPr.Strikeout)
						graphics.drawHorLine(0, strikeoutY, x0, x1, lineW);
					
					if (oNumTextPr.Underline)
						graphics.drawHorLine(0, underlineY, x0, x1, lineW);
				}
				
				if (true === editor.ShowParaMarks && (Asc.c_oAscNumberingSuff.Tab === nNumSuff || oNumLvl.IsLegacy()))
				{
					let tabSymbolWidth = 3.143; // ширина символа "стрелка влево" в шрифте Wingding3,10
					let tabX = this.X;
					
					if (!isRtl)
					{
						if (AscCommon.align_Left === nNumJc)
							tabX += numWidth;
						else if (AscCommon.align_Center === nNumJc)
							tabX += numWidth / 2;
					}
					
					graphics.SetFont({
						FontFamily : {Name : "ASCW3", Index : -1},
						FontSize   : 10,
						Italic     : false,
						Bold       : false
					});
					
					let tabCode = isRtl ? tab_Symbol_Rtl : tab_Symbol;
					if (suffWidth > tabSymbolWidth)
						graphics.FillText2(tabX + suffWidth / 2 - tabSymbolWidth / 2, Y, String.fromCharCode(tabCode), 0, suffWidth);
					else if (isRtl)
						graphics.FillText2(tabX, Y, String.fromCharCode(tabCode), 0, suffWidth);
					else
						graphics.FillText2(tabX, Y, String.fromCharCode(tabCode), tabSymbolWidth - suffWidth, suffWidth);
				}
			}
		}
		else if (para_PresentationNumbering === numItem.Type)
		{
			var bIsEmpty = para.IsEmpty();
			if ((para.LogicDocument && para.LogicDocument.IsVisioEditor()) || !bIsEmpty ||
				para.IsThisElementCurrent() ||
				para.Parent.IsSelectionUse() && para.Parent.IsSelectionEmpty() && para.Parent.Selection.StartPos === para.GetIndex())
			{
				if(!isRtl)
				{
					if (Pr.ParaPr.Ind.FirstLine < 0)
						numItem.Draw(X, Y, graphics, this);
					else
						numItem.Draw(para.Pages[CurPage].X + Pr.ParaPr.Ind.Left, Y, graphics, this);
				}
				else
				{
					if (Pr.ParaPr.Ind.FirstLine < 0)
						numItem.Draw(X + numItem.getVisibleWidth() / 2, Y, graphics, this);
					else
						numItem.Draw(para.Pages[CurPage].X + para.Pages[CurPage].XLimit - Pr.ParaPr.Ind.Right, Y, graphics, this);
				}
			}
		}
		
		this.X += numItem.getVisibleWidth();
	};
	ParagraphContentDrawState.prototype.calculateY = function(textPr)
	{
		let y = this.Y;

		// Superscript and subscript shift (Word's fixed factors, or PowerPoint's a:rPr/@baseline in slides)
		let scriptShift = textPr.GetScriptShift();
		if (scriptShift)
			y -= scriptShift * textPr.FontSize * g_dKoef_pt_to_mm;

		return y;
	};
	ParagraphContentDrawState.prototype.isSlideEditor = function()
	{
		return this.Paragraph && !this.Paragraph.bFromDocument;
	};
	ParagraphContentDrawState.prototype.isTextArtDraw = function()
	{
		return this.Graphics && this.Graphics.m_bIsTextDrawer;
	};
	ParagraphContentDrawState.prototype.isRtlMainDirection = function()
	{
		return this.Paragraph.isRtlDirection();
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.ParagraphContentDrawState = ParagraphContentDrawState;
	
})(window);

