// ##################################################################################################
// functionsEC regroupe toutes les fonctions utilisées dans les scripts développés par Emmanuel Côtez
// ##################################################################################################

// Fonction de création d'une fenêtre permettant la sélection d'un ou plusieurs éléments ;
// prend en paramètre un tableau d'occurrences
// renvoie un tableau contenant les réponses sélectionnées

// affiche le contenu d'un tableau passé en paramètre dans une boîte d'alerte
function alertTab(tab) {
	var msg="";
	for (i=0 ; i<tab.length ; i++) {
	msg += tab[i]+"\n";
	}
	msg = msg.replace(/###/g,"\t");
	alert(msg);
	}

// fonction permettant l'ajout et l'utilisation d'une couleur spécifique dans un document
 function myColorAdd(myDocument, myColorName, myColorModel, myColorValue){
    if(myColorValue instanceof Array == false){
        myColorValue = [(parseInt(myColorValue, 16) >> 16 ) & 0xff, (parseInt(myColorValue, 16) >> 8 ) & 0xff, parseInt(myColorValue, 16 ) & 0xff ];
        myColorSpace = ColorSpace.RGB;
    }else{
        if(myColorValue.length == 3)
          myColorSpace = ColorSpace.RGB;
        else
          myColorSpace = ColorSpace.CMYK;
    }
    try{
        myColor = myDocument.colors.item(myColorName);
        myName = myColor.name;
    }
    catch (myError){
        myColor = myDocument.colors.add();
        myColor.properties = {name:myColorName, model:myColorModel, space:myColorSpace ,colorValue:myColorValue};
    }
    return myColor;
}	

function myInput(type, msg, title, tabChoix) {

/*
var box = new Window('dialog', 'Some title');
var panel = box.add(panel, undefined, 'Panel title');
panel.add('edittext', undefined, 'Default value');
panel.add('slider', undefined, 50,0,100);
var group = box.add(group, undefined, 'Group title');
group = area_len_box.add('group', undefined, 'Title (not displayed)');
group.orientation='row';
group.closeBtn = group.add('button',undefined, 'Close', {name:'close'});
group.closeBtn.onClick = function(){
  box.hide();
  return false;
}
*/	

         var myWindow = new Window ("dialog", msg);
         
         var myInputGroup = myWindow.add ("group");

         // myInputGroup.add('slider', undefined, 50,0,100);

         if (type=="checkbox") {
         id=0;
         var check = new Array() ;
         myWindow.alignChildren = "left";
			
			for (j=0 ; j<tabChoix.length ; j++) {
			check[j] = myWindow.add ("checkbox", undefined, tabChoix[j]);
			// var check2 = myWindow.add ("checkbox", undefined, "Prefer black and white");
			}
			
			// check1.value = true;
		}

		if (type=="text") {
         
         myInputGroup.add ("statictext", undefined, title);
                  var myText = myInputGroup.add ("edittext", undefined, "John");
                  myText.characters = 20;
                  myText.active = true;
        }
            
     

 		var myButtonGroup = myWindow.add ("group");
                  myButtonGroup.alignment = "right";
                  myButtonGroup.add ("button", undefined, "OK");
                  myButtonGroup.add ("button", undefined, "Cancel");              

		if (myWindow.show () == 1) {
				
				if (type=="checkbox") {
					var tab = new Array() ;
					// alert (check[0].value) ;
					for (i=0 ; i<check.length ; i++) {
					if (check[i].value==true) {tab[i]="1";} else {tab[i]="0";}
					}
				return tab ; 
				}

				if (type=="text") { return myText.text ; }
              }
         		else {
                  exit ();
         		}

  } // fin de function selectWindow()


// cleanArray removes all duplicated elements
// pris sur : https://www.unicoda.com/?p=579
function cleanArray(array) {
  var i, j, len = array.length, out = [], obj = {};
  for (i = 0; i < len; i++) {
    obj[array[i]] = 0;
  }
  for (j in obj) {
    out.push(j);
  }
  return out;
}

// +++++++++++++++++++++++++
// Extrait une liste triée et dédoublonnée des intitulés de paragraphes dont le style est passé en paramètre "Matériel Diagnose"
// Fonction utilisée dans les scripts : test-paragraphs.js
// +++++++++++++++++++++++++

function getParagraphList() { 
	app.findGrepPreferences=app.changeGrepPreferences=null;
	app.findGrepPreferences.appliedParagraphStyle="Matériel-diagnose";
	// requête v1 : (?<=\r)[\w\(\) ~S~<-]+(?=\.\s—)
	// requête v2 (actuelle) : ^[\w\(\) ~S~<-]+(?=\.\s—)
    app.findGrepPreferences.findWhat="^[\\w\\(\\)\\[\\] ~S~<-]+(?=\\.\\s—)" ;
    
    res = app.activeDocument.findGrep();
        
    var tabResults = new Array();

    if (res.length!=0) {
	    for (i=0 ; i<res.length ; i++) {
	    	tabResults[i] = res[i].contents ;
	        }    
	    } else { return false ; }


	 // Suppression des doublons
	 tabResults = cleanArray(tabResults);

	// Tri du tableau
	 tabResults.sort();

	 return tabResults;
}



// Fonction de temporisation
function wait(ms)
{
	var d = new Date();
    var d2 = null;
    do { d2 = new Date(); }
    while(d2-d < ms);
}	

function remove_link(mySelectedText) {
		var myDoc = app.activeDocument;
		var mySel = app.selection[0];
		var  hyperlinkTextSources = mySelectedText.findHyperlinks(RangeSortOrder.ASCENDING_SORT);
		for (var ih = 0; ih < hyperlinkTextSources.length; ih++) {
			hyperlinkTextSources[ih].sourceText.appliedCharacterStyle = myDoc.characterStyles.itemByName("[Sans]"); 
			hyperlinkTextSources[ih].sourceText.clearOverrides(OverrideType.CHARACTER_ONLY);
			hyperlinkTextSources[ih].remove();
		}
	}	