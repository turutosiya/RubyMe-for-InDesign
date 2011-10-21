/*!
 * RubyMe for InDesign CS3+ v1.1
 * 
 * Copyright 2011,	Toshiya TSURU
 *							SunBusiness, Inc. : http://www.sunbic.co.jp/
 *
 * Date: Wed Oct 21 16:39:00 2011 +0900
 */
 /* ChangeLog * * * * 
	 
 2011-10-20  Toshiya TSURU < t_tsuru@sunbi.co.jp>
	* Fix Bugs
	* Add 'Ruby Auto Scaling' feature #45.

2011-10-21  Toshiya TSURU < t_tsuru@sunbi.co.jp>
	* Enhancement of Ruby Auto Scaling feature  
	
*/
// CONSTRAINTS 
var DELIMITER = '#';
var DEFAULT_RUBY_FONT_SIZE_BY_POINT = 4.606;

//
var doc = app.activeDocument;

// save current preferences
var pre_horizontalMeasurementUnits						= doc.viewPreferences.horizontalMeasurementUnits;
var pre_verticalMeasurementUnits							= doc.viewPreferences.verticalMeasurementUnits;
var pre_lineMeasurementUnits								= doc.viewPreferences.lineMeasurementUnits;
var pre_textSizeMeasurementUnits							= doc.viewPreferences.textSizeMeasurementUnits;
var pre_typographicMeasurementUnits					= doc.viewPreferences.typographicMeasurementUnits;
// unify the units!!
doc.viewPreferences.horizontalMeasurementUnits		=	MeasurementUnits.POINTS;
doc.viewPreferences.verticalMeasurementUnits			=	MeasurementUnits.POINTS;
doc.viewPreferences.lineMeasurementUnits				=	MeasurementUnits.POINTS;
doc.viewPreferences.textSizeMeasurementUnits		=	MeasurementUnits.POINTS;
doc.viewPreferences.typographicMeasurementUnits	=	MeasurementUnits.POINTS;

try{
	// var declaration
	var selections=doc.selection;
	// Exit if no selection
	for(var i = 0; i < selections.length; ++i ) {
		var selection	= selections[i];
		// continue to next selection
		if(!selection.contents) continue;
		if(selection.contents < 1) continue;
		
		// process the contents by unit of DELIMITERkanjiDELIMITERrubyDELIMITER
		for(var contents = selection.contents; -1 < contents.indexOf(DELIMITER, 0); contents = selection.contents) {
			// get indexOf DELIMITER
			var i1 = contents.indexOf(DELIMITER, i1);
			var i2 = contents.indexOf(DELIMITER, i1 + 1);
			var i3 = contents.indexOf(DELIMITER, i2 + 1);
			if (i1 < 0 || i2 < 0 || i3 < 0) break; // if not contains three '#', break !
			// left side -> kanji, right side -> ruby
			var kanji	= contents.substring(i1 + 1, i2);
			var ruby	= contents.substring(i2 + 1, i3);
			// replace string
			var target = selection.characters.itemByRange(i1, i3);
			target.contents = kanji;
			// set ruby
			target = selection.characters.itemByRange(i1, i1 + (kanji.length - 1));
			if(!target) continue;
			target.rubyType				= (1 < kanji.length) ? RubyTypes.GROUP_RUBY : RubyTypes.PER_CHARACTER_RUBY;
			target.rubyFlag				= true;
			target.rubyString			= ruby;
			// set ruby X Scale;
			var kanjiWidth	= selection.characters[i1 + (kanji.length - 1)].horizontalOffset - selection.characters[i1].horizontalOffset + selection.characters[i1 + (kanji.length - 1)].pointSize;
			var rubyWidth	= ruby.length * ((target.rubyFontSize < 0) ? DEFAULT_RUBY_FONT_SIZE_BY_POINT : target.rubyFontSize);
			if(kanjiWidth < rubyWidth) target.rubyXScale = (kanjiWidth * 100) / rubyWidth;
			else target.rubyXScale	= 100;
		}
	}
}finally{
	doc.viewPreferences.horizontalMeasurementUnits		=	pre_horizontalMeasurementUnits;
	doc.viewPreferences.verticalMeasurementUnits			=	pre_verticalMeasurementUnits;
	doc.viewPreferences.lineMeasurementUnits				=	pre_lineMeasurementUnits;
	doc.viewPreferences.textSizeMeasurementUnits		=	pre_textSizeMeasurementUnits;
	doc.viewPreferences.typographicMeasurementUnits	=	pre_typographicMeasurementUnits;
}