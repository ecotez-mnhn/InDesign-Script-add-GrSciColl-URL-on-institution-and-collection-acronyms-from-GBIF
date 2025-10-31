// ———————————————————————————————————————————————————————————————————————————————————————————————
// InDesign script to get GBIF GrSciColl ID from cited acronyms in the text 
// ———————————————————————————————————————————————————————————————————————————————————————————————
// Script name: "get_GBIF_GrSciColl_IDs.jsx"
// ———————————————————————————————————————————————————————————————————————————————————————————————
// Written by: Emmanuel Côtez (emmanuel.cotez@mnhn.fr, emmanuel.cotez@teznet.fr)
// ———————————————————————————————————————————————————————————————————————————————————————————————
// /!\ Read the read.me file for more details on how to install the script and how it works /!\
// ———————————————————————————————————————————————————————————————————————————————————————————————
// Available on gitHub at the address: https://github.com/ecotez-mnhn
// ———————————————————————————————————————————————————————————————————————————————————————————————
// v.1.01 [31 October 2025, "Halloween update"]
// ———————————————————————————————————————————————————————————————————————————————————————————————

// Beginning of the script

// Dependencies
#include 'functionsEC.js';
#include 'extendables/extendables.jsx';
#include 'restix.jsx';

// Variables
var count=0 ;

// Name of the character style that will be applied on created links (by default "Lien hypertexte")
var appliedHyperlinkCharacterStyle="Lien hypertexte";
// var appliedHyperlinkCharacterStyle="your_character_style_name";

// Name of the paragraph style that contains the acronyms to link (by default "Abréviations")
var searchedAbbreviationParagraphStyle="Abréviations";
// var searchedAbbreviationParagraphStyle="your_paragraph_style_name";

// Functions
function style_exist(style_name) {
	// test l'existence d'un style et le créé s'il n'existe pas
	if(app.activeDocument.characterStyles.item (style_name) == null) { 
		app.activeDocument.characterStyles.add({name:style_name});
	}	
	return app.activeDocument.characterStyles.item(style_name);
}

function trim (str) {
	// function written by Peter Kahrel
	// available there: https://community.adobe.com/t5/indesign-discussions/trim-function-not-working-in-extendscript-is-there-a-way-to-replicate-the-functioning/td-p/7671885
    return str.replace(/^\s+/,'').replace(/\s+$/,'');
	}
	
function copy2clipboard (s) {
    var tf = app.activeDocument.pages[0].textFrames.add()
	tf.parentStory.texts[0].select();
    tf.parentStory.contents = s;
    tf.parentStory.texts.everyItem().select(SelectionOptions.REPLACE_WITH)
	app.copy()
    tf.remove()
}	

myColorAdd(app.activeDocument, appliedHyperlinkCharacterStyle, ColorModel.PROCESS, [100,90,10,0]); // hyperlink color
myColorAdd(app.activeDocument, "SCRIPT-COLOR-GrSciColl-red-error", ColorModel.PROCESS, [255,0,0]); // red (error) 

myStyleHyperlink = style_exist(appliedHyperlinkCharacterStyle);
myStyleHyperlink.fillColor = appliedHyperlinkCharacterStyle;

var cont = true ;

var mySel = app.selection[0];

 if (mySel== null || mySel.contents.length == 0) {
	 // if no active selection, search the paragraph style: searchedAbbreviationParagraphStyle
	if ( app.activeDocument.paragraphStyles.everyItem().name.join().indexOf(searchedAbbreviationParagraphStyle) == -1) {
		alert("The paragraph style '"+searchedAbbreviationParagraphStyle+"' does not exist; nothing to search.\n\n—> You can configure your own paragraph style name in the script.");
		exit();
		}

	app.findGrepPreferences=app.changeGrepPreferences=null;
	app.findGrepPreferences.findWhat="^.+(?=\\t)";
	app.findGrepPreferences.appliedParagraphStyle=searchedAbbreviationParagraphStyle;
	app.findGrepPreferences.appliedCharacterStyle="[No character style]";
	acronyms = app.activeDocument.findGrep(); 	 
	cont = false ; 

	 } else if (mySel.contents.length > 10) { 
			alert ("Please select an acronym only (e.g., \"P\")");
			exit();
			}
		else {
		app.findGrepPreferences=app.changeGrepPreferences=null;
		app.findGrepPreferences.findWhat=".+";
		acronyms = mySel.findGrep();
	 }  

// Looking for GrSciColl hyperlinks with GBIF API

for (i=0; i<acronyms.length;i++) {
	var acroStr = trim(acronyms[i].contents) ;
	acronyms[i].select();
	acronyms[i].showText();
	
	var url_type = "" ;
	
	// GBIF API request
	var yourUrl = "http://api.gbif.org/v1/grscicoll/institution/search?code="+acroStr ;
	request = {
	  url:yourUrl,
	  }
	var response = restix.fetch(request);

	try { var json_obj = JSON.parse(response.body); } catch (e) { }
	
	// multiple answers	
	var noms_complets = new Array() ;
	var acros_complets = new Array() ;
	var jmax = json_obj["count"];
	for (j=0;j<jmax;j++) {
		noms_complets[j] = json_obj["results"][j]["name"];
		acros_complets[j] = json_obj["results"][j]["code"];
		} // end multiple answers

	var acronym_ok_id = "" ;
	
	var clipboard = "" ;

	if (jmax>0) {
		var url_type = "institution" ;
		// If at least 1 result
		if (jmax==1) {
			// If only 1 result
			if (json_obj["results"][0]["code"]==acroStr) {
				// Checking if the acronym corresponds to the sent result
				var req = confirm("Please confirm found institution:\n"+json_obj["results"][0]["name"]);
				if (req==false) { acronym_ok_id=""; }
				else {
					acronym_ok_id = json_obj["results"][0]["key"];
					// Copying the acronym name in the clipboard
					clipboard = json_obj["results"][0]["name"];
					}
			}
		} else if (jmax>1) {
			// If more than only 1 result, asking which one should be chosen
			// Choosing the right result in a list
			var msg="";
			 var myCaseList = new Array();
			 var myCase ;
			 for (m=0 ; m<noms_complets.length ; m++) {
				myCaseList[m] = acros_complets[m] + " / " + noms_complets[m] ;
				msg = msg + noms_complets[m] + " / " ;
			 }

			 var myDialog = app.dialogs.add({name:"There are more than one result for the acronym [Institution "+acroStr+"]:"})
			 with(myDialog){
			  with(dialogColumns.add()){
			   with (dialogRows.add()) {
				with (dialogColumns.add()) {
				 staticTexts.add({staticLabel:"Which one should be linked?"});
				}
				with (dialogColumns.add()) {
				 myCase = dropdowns.add({stringList:myCaseList,selectedIndex:0,minWidth:200});
				}
			   }
			  }
			 }
			 var myResult = myDialog.show();

			 if (myResult != true){
			  // user clicked Cancel
			  acronym_ok_id="";
			  myDialog.destroy();
			 } else {
			  theCase = myCase.selectedIndex;
			  acronym_ok_id = json_obj["results"][theCase]["key"];
			  clipboard = json_obj["results"][theCase]["name"];
			  myDialog.destroy();
			  // Result was selected by user
			 }
		} 
	} 
	
	if (acronym_ok_id=="") {
		
		// alert ("Institution acronym ID not found, searching collection ID");
		
		var url_type = "collection" ;
		
		// If no result, searching for *collection* acronyms
		
			// request model on the GBIF : https://api.gbif.org/v1/grscicoll/institution/suggest?q=ANSP
	
			var yourUrl = "http://api.gbif.org/v1/grscicoll/collection/search?code="+acroStr ;
						
			request = {
			  url:yourUrl,
			  }
			var response = restix.fetch(request);

			try { var json_obj = JSON.parse(response.body); } catch (e) {}
			
			// multiple answers	
			var noms_complets = new Array() ;
			var acros_complets = new Array() ;
			var jmax = json_obj["count"];
			for (j=0;j<jmax;j++) {
				noms_complets[j] = json_obj["results"][j]["name"];
				acros_complets[j] = json_obj["results"][j]["code"];
				} // end multiple answers

			var acronym_ok_id = "" ;

			if (jmax>0) {
				// If at least 1 result
				if (jmax==1) {
					// If only 1 result
					if (json_obj["results"][0]["code"]==acroStr) {
						// Checking if the acronym corresponds to the sent result
						var req = confirm("Please confirm found collection:\n"+json_obj["results"][0]["name"]);
						if (req==false) { acronym_ok_id=""; }
						else {
							acronym_ok_id=json_obj["results"][0]["key"];
							// Copying the acronym name in the clipboard
							clipboard = json_obj["results"][0]["name"];							
							}
						
					}
				} else if (jmax>1) {
					// If more than only 1 result, asking which one should be chosen
					// Choosing the right result in a list
					var msg="";
					 var myCaseList = new Array();
					 var myCase ;
					 for (m=0 ; m<noms_complets.length ; m++) {
						myCaseList[m] = acros_complets[m] + " / " + noms_complets[m] ;
						msg = msg + noms_complets[m] + " / " ;
					 }

					 var myDialog = app.dialogs.add({name:"There are more than one result for the acronym [Collection "+acroStr+"]:"})
					 with(myDialog){
					  with(dialogColumns.add()){
					   with (dialogRows.add()) {
						with (dialogColumns.add()) {
						 staticTexts.add({staticLabel:"Which one should be linked?"});
						}
						with (dialogColumns.add()) {
						 myCase = dropdowns.add({stringList:myCaseList,selectedIndex:0,minWidth:200});
						}
					   }
					  }
					 }
					 var myResult = myDialog.show();

					 if (myResult != true){
					  // user clicked "Cancel"
					  acronyms[i].strokeColor="SCRIPT-COLOR-GrSciColl-red-error";	
					  acronym_ok_id="";
					  myDialog.destroy();
					 } else {
					  theCase = myCase.selectedIndex;
					  acronym_ok_id = json_obj["results"][theCase]["key"];
					  clipboard = json_obj["results"][theCase]["name"];
					  myDialog.destroy();
					  
					  // Result was selected by user
					  
					 }
					// #############################################
				}
			} else {
				// alert ("xxx") ;
				acronyms[i].strokeColor="SCRIPT-COLOR-GrSciColl-red-error";	
				acronym_ok_id="";	
			}
		// @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@		
		}


	if (acronym_ok_id!="") { 
	
		// Copy the actual acronym name in the clipboard
		copy2clipboard(clipboard);
		
		// acronym_ok_id contains the ID of the institution or collection, validated by user
		
		acronyms[i].select();
		mySelectedText=app.selection[0];
		remove_link(mySelectedText);
		mySelectedText.strokeColor="None";
		
		if (url_type=="institution") {
			var url = "https://registry.gbif.org/institution/"+acronym_ok_id ;
		} else {
			var url = "https://registry.gbif.org/collection/"+acronym_ok_id ;
		}
		try {
		var myHyperlinkSource = app.activeDocument.hyperlinkTextSources.add(acronyms[i]) ;
		var myHyperlinkURLDestination = app.activeDocument.hyperlinkURLDestinations.add(url) ;
	    var myHyperlink = app.activeDocument.hyperlinks.add(myHyperlinkSource, myHyperlinkURLDestination, {name: "GrSciColl_"+acroStr+" ("+(count)+")"}) ; 
		acronyms[i].appliedCharacterStyle=appliedHyperlinkCharacterStyle;
	    count++;
		url_type="";
		} catch (e) {
		// alert ("error when creating hyperlink on "+clipboard);
		acronyms[i].strokeColor="SCRIPT-COLOR-GrSciColl-red-error";	
		}
	}
}
// End of the script