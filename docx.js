
/*
 * DOCX.js - Generate .docx files using pure client-side JavaScript.
 */

var DOCXjs = function() {
	
	var parts = {};

	// Content store
	var textElements = [];

	/* This is the file that sits in the root of the DOCX file, and specifies the mimetypes of the included files */
	var contentTypes = function() {
		var output = '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>';
		output += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
		
		// Add defaults
		output += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"> </Default>';
		output += '<Default Extension="xml" ContentType="application/xml"> </Default>'
		
		// Overrides
		var overrides = {
			'/word/numbering.xml': 'wordprocessingml.numbering',
			'/word/styles.xml': 'wordprocessingml.styles',
			'/docProps/app.xml': 'extended-properties',
			'/word/settings.xml': 'wordprocessingml.settings',
			'/word/theme/theme1.xml': 'theme',
			'/word/fontTable.xml': 'fontTable',
			'/word/webSettings.xml': 'webSettings',
			'/docProps/core.xml': 'core-properties',
			'/word/document.xml': 'document.main'
		}
		
		for (var override in overrides) {
			output += '<Override PartName="' + override + '" ContentType="application/vnd.openxmlformats-officedocument.' + overrides[override] + '+xml"></Override>';
		}
		
		output += '</Types>';
		
		return output;
	}
	
	var documentGen = function() {
		var output = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
		
		// Word document element with XML namespaces
		output += '<w:wordDocument ';
		output += ' xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006"'
		output += ' xmlns:o="urn:schemas-microsoft-com:office:office"';
		output += ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
		output += ' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
		output += ' xmlns:v="urn:schemas-microsoft-com:vml"';
		output += ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
		output += ' xmlns:w10="urn:schemas-microsoft-com:office:word" ';
		output += ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
		output += ' xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml">';
		
		// Main body of the document
		output += '<w:body>';
		
		for (var textElement in textElements) {
			// Start a text section
			output += '<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"></w:rFonts><w:b w:val="single"></w:b></w:rPr>';
			
			// @TODO: Do escaping
			output += '<w:t xml:space="preserve">' + textElements[textElement] + '</w:t></w:r></w:p>';
		}
		
		
		// Section 'sign-off', dimensions, gutter etc.
		output +='<w:sectPr w:rsidR="12240" w:rsidRPr="12240" w:rsidSect="12240"><w:pgSz w:w="11906" w:h="16838"></w:pgSz><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="0" w:footer="0" w:gutter="0"></w:pgMar><w:cols w:space="708"></w:cols><w:docGrid w:linePitch="360"></w:docGrid></w:sectPr>';
		
		// Close
		output += '</w:body></w:wordDocument>';
		
		return output;
	}
	
	
	var generate = function() {
		// Content types
		
		var files = [
			'[Content_Types].xml',
			'docProps/app.xml',
			'docProps/core.xml',
			'word/_rels/document.xml.rels',
			'word/document.xml',
			'word/endnotes.xml',
			'word/fontTable.xml',
			'word/footer1.xml',
			'word/footer2.xml',
			'word/footer3.xml',
			'word/footnotes.xml',
			'word/header1.xml',
			'word/header2.xml',
			'word/header3.xml',
			'word/settings.xml',
			'word/styles.xml',
			'word/webSettings.xml',
			'word/theme/theme1.xml'
		];
		
		var file_data = {};
		
		var file_count = files.length;
		var file_count_current = 0;
		var zip = new JSZip("STORE");
		zip.folder('_rels');
		
		for(var file in files) {
			$.ajax({
				url: 'blank/' + files[file],
				complete: function(r) {
					//file_data[this.url.replace(/blank_/, '')] = r.responseText;
					
					zip.add(this.url.replace('blank/', ''), r.responseText);
					file_count_current ++;
					
					if (file_count == file_count_current) {
						console.log('DONE!');
						outputFile = zip.generate();
						document.location.href = 'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' + outputFile;
					}
					
				}
			});
		}		
		
		return parts;
		
	}
	
	
	// Add content methods
	
	var addText = function(string) {
		textElements.push(string);
	}
	
	var finalFile = function(parts) {
		var zip = new JSZip();
		for (var part in parts) {
			zip.add(part, parts[part]);
		}
		return zip.generate();
	}
	
	return {
		output: function(type, options) {
			
			var buffer = generate();
			return;
			if(type == undefined) {
				
				return buffer;
			}
			if(type == 'datauri') {
				var outputFile = finalFile(buffer);
				document.location.href = 'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' + outputFile;
			}
		// @TODO: Add different output options
		},
		text: function(string) {
			addText(string);
		}
	};
	
};