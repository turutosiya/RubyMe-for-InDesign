/*!
 * RubyMe for InDesign CS3+ v1.1
 * 
 * Copyright 2011,	Toshiya TSURU
 *							SunBusiness, Inc. : http://www.sunbic.co.jp/
 *
 * Date: Wed Oct 20 10:07:00 2011 +0900
 */
 /* ChangeLog * * * * 
	 
 2011-10-20  Toshiya TSURU < t_tsuru@sunbi.co.jp>
	* Fix Bugs
	* Add 'Ruby Auto Scaling' feature #45.

*/
// 
var DELIMITER = '#';

// var declaration
var selections=app.activeDocument.selection;
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
		target.contents		= kanji;
		// set ruby
		target = selection.characters.itemByRange(i1, i1 + (kanji.length - 1));
		if(!target) continue;
		if(1 < kanji.length) target.rubyType		= RubyTypes.GROUP_RUBY;
		target.rubyString			= ruby;
		target.rubyFlag				= true;
		target.rubyAutoScaling	= true;
	}
}