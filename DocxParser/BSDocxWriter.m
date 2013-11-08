//
//  BSDocxWriter.m
//  DocxParser
//
//  Created by Bradley Slayter on 11/7/13.
//  Copyright (c) 2013 Brad Slayter. All rights reserved.
//

#import "BSDocxWriter.h"

@interface NSAttributedString (RangeExtension)
- (NSRange)fullRange;
@end

@implementation NSAttributedString (RangeExtension)
- (NSRange)fullRange {
	return (NSRange){0, self.string.length};
}
@end

@implementation BSDocxWriter

-(id) initWithAttributedString:(NSAttributedString *)string {
	if ((self = [super init])) {
		self.string = string;
		
		[self splitString];
	}
	return self;
}

-(NSData *) buildDocument {
	GDataXMLElement *root = [self buildRootElement];
	GDataXMLElement *bodyElem = [self buildBody];
	[root addChild:bodyElem];
	
	self.xmlDocument = [[GDataXMLDocument alloc] initWithRootElement:root];
	return self.xmlDocument.XMLData;
}

-(void) splitString {
	NSMutableArray *tempParagraphs = [[NSMutableArray alloc] init];
	
	// Split string into substrings by similar attributes
	[self.string enumerateAttributesInRange:[self.string fullRange] options:0 usingBlock:^(NSDictionary *attrs, NSRange range, BOOL *stop) {
		
		NSString *subString = [[self.string string] substringWithRange:range];
		NSAttributedString *attrString = [[NSAttributedString alloc] initWithString:subString attributes:attrs];
		
		[tempParagraphs addObject:attrString];
									 
	}];
	
	paragraphs = [[NSMutableArray alloc] init];
	for (NSAttributedString *subString in tempParagraphs) {
		// split those strings into smaller strings by newline char
		NSArray *subSubStrings = [self splitStringByNewLine:subString.string];
		//NSLog(@"SUB: %@", subSubStrings);
		
		for (NSString *s in subSubStrings) {
			// re-build substrings into attributed strings
			NSAttributedString *attrString = [[NSAttributedString alloc] initWithString:s attributes:[subString attributesAtIndex:0 effectiveRange:NULL]];
			[paragraphs addObject:attrString];
			/*if ([s isEqual:[subSubStrings lastObject]]) {
				NSAttributedString *newLineParagraph = [[NSAttributedString alloc] initWithString:@"\n" attributes:[subString attributesAtIndex:0 effectiveRange:NULL]];
				[paragraphs addObject:newLineParagraph];
			}*/
		}
	}
	
	NSMutableArray *tempArray = [[NSMutableArray alloc] init];;
	tempParagraphs = [[NSMutableArray alloc] initWithArray:paragraphs];
	[paragraphs removeAllObjects];
	for (NSAttributedString *subString in tempParagraphs) {
		[tempArray addObject:subString];
		
		if ([subString.string hasSuffix:@"\n"]) {
			NSAttributedString *newSubString = [[NSAttributedString alloc] initWithString:[subString.string stringByReplacingOccurrencesOfString:@"\n" withString:@""] attributes:[subString attributesAtIndex:0 effectiveRange:nil]];
			[tempArray removeObject:subString];
			[tempArray addObject:newSubString];
			
			NSArray *a = [[NSArray alloc] initWithArray:tempArray];
			[paragraphs addObject:a];
			[tempArray removeAllObjects];
		}
	}
	
	NSLog(@"Printing: %d", (int)paragraphs.count);
	for (NSMutableArray *a in paragraphs) {
		for (NSAttributedString *s in a)
			NSLog(@"%@", [s.string stringByReplacingOccurrencesOfString:@"\n" withString:@"\\n"]);
	}
	
	[self buildRootElement];
}

-(NSArray *) splitStringByNewLine:(NSString *)string {
	NSMutableArray *lines = [NSMutableArray array];
	NSRange searchRange = NSMakeRange(0, string.length);
	while (1) {
		NSRange newLineRange = [string rangeOfString:@"\n" options:NSLiteralSearch range:searchRange];
		if (newLineRange.location != NSNotFound) {
			NSInteger index = newLineRange.location + newLineRange.length;
			NSString *line = [string substringWithRange:NSMakeRange(searchRange.location, index - searchRange.location)];
			[lines addObject:line];
			searchRange = NSMakeRange(index, string.length - index);
		} else {
			break;
		}
	}
	
	return lines;
}

-(GDataXMLElement *) buildRootElement {
	GDataXMLElement *root = [GDataXMLElement elementWithName:@"w:document"];
	
	GDataXMLNode *wpc = [GDataXMLNode elementWithName:@"xmlns:wpc" stringValue:@"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"];
	[root addAttribute:wpc];
	
	GDataXMLNode *mc = [GDataXMLNode elementWithName:@"xmlns:mc" stringValue:@"http://schemas.openxmlformats.org/markup-compatibility/2006"];
	[root addAttribute:mc];
	
	GDataXMLNode *o = [GDataXMLNode elementWithName:@"xmlns:o" stringValue:@"urn:schemas-microsoft-com:office:office"];
	[root addAttribute:o];
	
	GDataXMLNode *r = [GDataXMLNode elementWithName:@"xmlns:r" stringValue:@"http://schemas.openxmlformats.org/officeDocument/2006/relationships"];
	[root addAttribute:r];
	
	GDataXMLNode *m = [GDataXMLNode elementWithName:@"xmlns:m" stringValue:@"http://schemas.openxmlformats.org/officeDocument/2006/math"];
	[root addAttribute:m];
	
	GDataXMLNode *v = [GDataXMLNode elementWithName:@"xmlns:v" stringValue:@"urn:schemas-microsoft-com:vml"];
	[root addAttribute:v];
	
	GDataXMLNode *wp14 = [GDataXMLNode elementWithName:@"xmlns:wp14" stringValue:@"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"];
	[root addAttribute:wp14];
	
	GDataXMLNode *wp = [GDataXMLNode elementWithName:@"xmlns:wp" stringValue:@"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"];
	[root addAttribute:wp];
	
	GDataXMLNode *wp10 = [GDataXMLNode elementWithName:@"xmlns:wp10" stringValue:@"urn:schemas-microsoft-com:office:word"];
	[root addAttribute:wp10];
	
	GDataXMLNode *w = [GDataXMLNode elementWithName:@"xmlns:w" stringValue:@"http://schemas.openxmlformats.org/wordprocessingml/2006/main"];
	[root addAttribute:w];
	
	GDataXMLNode *w14 = [GDataXMLNode elementWithName:@"xmlns:w14" stringValue:@"http://schemas.microsoft.com/office/word/2010/wordml"];
	[root addAttribute:w14];
	
	GDataXMLNode *w15 = [GDataXMLNode elementWithName:@"xmlns:w15" stringValue:@"http://schemas.microsoft.com/office/word/2012/wordml"];
	[root addAttribute:w15];
		
	GDataXMLNode *wpg = [GDataXMLNode elementWithName:@"xmlns:wpg" stringValue:@"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"];
	[root addAttribute:wpg];
	
	GDataXMLNode *wpi = [GDataXMLNode elementWithName:@"xmlns:wpi" stringValue:@"http://schemas.microsoft.com/office/word/2010/wordprocessingInk"];
	[root addAttribute:wpi];
	
	GDataXMLNode *wne = [GDataXMLNode elementWithName:@"xmlns:wne" stringValue:@"http://schemas.microsoft.com/office/word/2006/wordml"];
	[root addAttribute:wne];
	
	GDataXMLNode *wps = [GDataXMLNode elementWithName:@"xmlns:wps" stringValue:@"http://schemas.microsoft.com/office/word/2010/wordprocessingShape"];
	[root addAttribute:wps];
	
	GDataXMLNode *mcI = [GDataXMLNode elementWithName:@"mc:Ignorable" stringValue:@"w14 w15 wp14"];
	[root addAttribute:mcI];
	
	return root;
}

-(GDataXMLElement *) buildBody {
	GDataXMLElement *bodyElement = [GDataXMLElement elementWithName:@"w:body"];
	
	for (NSArray *paragraph in paragraphs) {
		GDataXMLElement *paragraphElement = [GDataXMLElement elementWithName:@"w:p"];
		
		GDataXMLNode *rsidR = [GDataXMLNode elementWithName:@"w:rsidR" stringValue:@"00000000"];
		[paragraphElement addAttribute:rsidR];
		GDataXMLNode *rsidRDefault = [GDataXMLNode elementWithName:@"w:rsidRDefault" stringValue:@"00000000"];
		[paragraphElement addAttribute:rsidRDefault];
		GDataXMLNode *rsidP = [GDataXMLNode elementWithName:@"w:rsidP" stringValue:@"00000000"];
		[paragraphElement addAttribute:rsidP];
		
		// check alignment
		NSAttributedString *firstRun;
		for (NSAttributedString *string in paragraph) { // get first non nil string in paragraph
			if (string.length)
				firstRun = string;
		}
		
		// read attributes
		NSDictionary *attrs = [firstRun attributesAtIndex:0 effectiveRange:nil];
		NSParagraphStyle *style = attrs[NSParagraphStyleAttributeName];
		// if alignment, add it to pPr
		if (style) {
			GDataXMLElement *alignElement = [self getAlignmentElement:style];
			GDataXMLElement *pPr = [GDataXMLElement elementWithName:@"w:pPr"];
			[pPr addChild:alignElement];
			[paragraphElement addChild:pPr];
		}
		
		for (NSAttributedString *runString in paragraph) {
			GDataXMLElement *runElement = [GDataXMLElement elementWithName:@"w:r"];
			[runElement addAttribute:rsidR];
			GDataXMLElement *runPr = [GDataXMLElement elementWithName:@"w:rPr"];
			
			// get string attributes
			NSDictionary *attrs;
			if (runString.length)
				attrs = [runString attributesAtIndex:0 effectiveRange:nil];
			
			// get font object
			UIFont *font = attrs[NSFontAttributeName];
			if (font) {
				// font size
				if (font.pointSize) {
					GDataXMLElement *sizeElement = [self getFontSizeElement:font];
					[runPr addChild:sizeElement];
				}
				
				// font name
				GDataXMLElement *fontNameElement = [self fontFamilyNameForFont:font];
				[runPr addChild:fontNameElement];
				
				// bold
				GDataXMLElement *boldElement = [self getBoldForFont:font];
				if (boldElement)
					[runPr addChild:boldElement];
				
				// italic
				GDataXMLElement *italicElement = [self getItalicForFont:font];
				if (italicElement)
					[runPr addChild:italicElement];
			}
			
			// get underline value
			NSNumber *underLineVal = attrs[NSUnderlineStyleAttributeName];
			if (underLineVal && underLineVal.intValue == NSUnderlineStyleSingle) {
				[runPr addChild:[self getUnderlineElement]];
			}
			
			// get font color
			UIColor *color = attrs[NSForegroundColorAttributeName];
			if (color) {
				GDataXMLElement *colorElement = [self getFontColorElement:color];
				[runPr addChild:colorElement];
			}
			
			// get actual text
			GDataXMLElement *textElement = [GDataXMLElement elementWithName:@"w:t" stringValue:runString.string];
			GDataXMLNode *spaceNode = [GDataXMLNode elementWithName:@"xml:space" stringValue:@"preserve"];
			[textElement addAttribute:spaceNode];
			if ([runString.string isEqualToString:@""]) {
				continue;
			}
			
			// add children
			[runElement addChild:runPr];
			[runElement addChild:textElement];
			[paragraphElement addChild:runElement];
		}
		
		[bodyElement addChild:paragraphElement];
	}
	
	[bodyElement addChild:[self buildSectPrElement]];
	
	return bodyElement;
}

-(GDataXMLElement *) buildSectPrElement {
	GDataXMLElement *sectPrElement = [GDataXMLElement elementWithName:@"w:sectPr"];
	
	GDataXMLNode *rsidR = [GDataXMLNode elementWithName:@"w:rsidR" stringValue:@"00000000"];
	GDataXMLNode *rsidRPr = [GDataXMLNode elementWithName:@"w:rsidRPr" stringValue:@"00000000"];
	[sectPrElement addAttribute:rsidR];
	[sectPrElement addAttribute:rsidRPr];
	
	GDataXMLElement *pgSz = [GDataXMLElement elementWithName:@"w:pgSz"];
	GDataXMLNode *w = [GDataXMLNode elementWithName:@"w:w" stringValue:@"12240"];
	GDataXMLNode *h = [GDataXMLNode elementWithName:@"w:h" stringValue:@"15840"];
	[pgSz addAttribute:w];
	[pgSz addAttribute:h];
	
	GDataXMLElement *pgMar = [GDataXMLElement elementWithName:@"w:pgMar"];
	GDataXMLNode *top = [GDataXMLNode elementWithName:@"w:top" stringValue:@"1440"];
	GDataXMLNode *right = [GDataXMLNode elementWithName:@"w:right" stringValue:@"1440"];
	GDataXMLNode *bottom = [GDataXMLNode elementWithName:@"w:bottom" stringValue:@"1440"];
	GDataXMLNode *left = [GDataXMLNode elementWithName:@"w:left" stringValue:@"1440"];
	GDataXMLNode *header = [GDataXMLNode elementWithName:@"w:header" stringValue:@"720"];
	GDataXMLNode *footer = [GDataXMLNode elementWithName:@"w:footer" stringValue:@"720"];
	GDataXMLNode *gutter = [GDataXMLNode elementWithName:@"w:left" stringValue:@"0"];
	[pgMar addAttribute:top];
	[pgMar addAttribute:right];
	[pgMar addAttribute:bottom];
	[pgMar addAttribute:left];
	[pgMar addAttribute:header];
	[pgMar addAttribute:footer];
	[pgMar addAttribute:gutter];
	
	GDataXMLElement *cols = [GDataXMLElement elementWithName:@"w:cols"];
	GDataXMLNode *space = [GDataXMLNode elementWithName:@"w:space" stringValue:@"720"];
	[cols addAttribute:space];
	
	GDataXMLElement *docGrid = [GDataXMLElement elementWithName:@"w:docGrid"];
	GDataXMLNode *linePitch = [GDataXMLNode elementWithName:@"w:linePitch" stringValue:@"360"];
	[docGrid addAttribute:linePitch];
	
	[sectPrElement addChild:pgSz];
	[sectPrElement addChild:pgMar];
	[sectPrElement addChild:cols];
	[sectPrElement addChild:docGrid];
	
	return sectPrElement;
}

-(GDataXMLElement *) getAlignmentElement:(NSParagraphStyle *)paragraphStyle {
	GDataXMLElement *alignElement = [GDataXMLElement elementWithName:@"w:jc"];
	GDataXMLNode *alignNode = [GDataXMLNode elementWithName:@"w:val"];
	
	switch (paragraphStyle.alignment) {
		case NSTextAlignmentCenter:
			[alignNode setStringValue:@"center"];
			break;
		case NSTextAlignmentRight:
			[alignNode setStringValue:@"right"];
			break;
		case NSTextAlignmentLeft:
			[alignNode setStringValue:@"left"];
			break;
		case NSTextAlignmentJustified:
			[alignNode setStringValue:@"justified"];
			break;
		default:
			break;
	}
	
	[alignElement addAttribute:alignNode];
	
	return alignElement;
}

-(GDataXMLElement *) getFontSizeElement:(UIFont *)font {
	GDataXMLElement *sizeElement = [GDataXMLElement elementWithName:@"w:sz"];
	GDataXMLNode *sizeNode = [GDataXMLNode elementWithName:@"w:val" stringValue:[NSString stringWithFormat:@"%d", (int)font.pointSize*2]];
	[sizeElement addAttribute:sizeNode];
	
	return sizeElement;
}

-(GDataXMLElement *) getBoldForFont:(UIFont *)font {
	if ([font.fontName rangeOfString:@"bold" options:NSCaseInsensitiveSearch].location != NSNotFound) {
		GDataXMLElement *boldElement = [GDataXMLElement elementWithName:@"w:b"];
		GDataXMLNode *boldVal = [GDataXMLNode elementWithName:@"w:val" stringValue:@"1"];
		[boldElement addAttribute:boldVal];
		
		return boldElement;
	} else {
		return nil;
	}
}

-(GDataXMLElement *) getItalicForFont:(UIFont *)font {
	if ([font.fontName rangeOfString:@"italic" options:NSCaseInsensitiveSearch].location != NSNotFound) {
		GDataXMLElement *italicElement = [GDataXMLElement elementWithName:@"w:i"];
		GDataXMLNode *italicVal = [GDataXMLNode elementWithName:@"w:val" stringValue:@"1"];
		[italicElement addAttribute:italicVal];
		
		return italicElement;
	} else {
		return nil;
	}
}

-(GDataXMLElement *) getUnderlineElement {
	NSLog(@"*******UNDERLINE");
	GDataXMLElement *underlineElement = [GDataXMLElement elementWithName:@"w:u"];
	GDataXMLNode *underlineNode = [GDataXMLNode elementWithName:@"w:val" stringValue:@"single"];
	[underlineElement addAttribute:underlineNode];
	
	return underlineElement;
}

-(GDataXMLElement *) getFontColorElement:(UIColor *)color {
	NSString *colorString = [self getHexStringForColor:color];
	GDataXMLElement *colorElement = [GDataXMLElement elementWithName:@"w:color"];
	GDataXMLNode *colorNode = [GDataXMLNode elementWithName:@"w:val" stringValue:colorString];
	[colorElement addAttribute:colorNode];
	
	return colorElement;
}

-(GDataXMLElement *) fontFamilyNameForFont:(UIFont *)font {
	
	GDataXMLElement *fontNameElement = [GDataXMLElement elementWithName:@"w:rFonts"];
	GDataXMLNode *csNode = [GDataXMLNode elementWithName:@"w:cs" stringValue:font.familyName];
	GDataXMLNode *asciiNode = [GDataXMLNode elementWithName:@"w:ascii" stringValue:font.familyName];
	GDataXMLNode *hAnsiNode = [GDataXMLNode elementWithName:@"w:hAnsi" stringValue:font.familyName];
	GDataXMLNode *eastAsiaNode = [GDataXMLNode elementWithName:@"w:eastAsia" stringValue:font.familyName];
	[fontNameElement addAttribute:csNode];
	[fontNameElement addAttribute:asciiNode];
	[fontNameElement addAttribute:hAnsiNode];
	[fontNameElement addAttribute:eastAsiaNode];
	
	return fontNameElement;
	
}

-(NSString *) getHexStringForColor:(UIColor *)color {
	const CGFloat *components = CGColorGetComponents(color.CGColor);
	CGFloat r = components[0];
	CGFloat g = components[1];
	CGFloat b = components[2];
	NSString *hexString=[NSString stringWithFormat:@"%02X%02X%02X", (int)(r * 255), (int)(g * 255), (int)(b * 255)];
	return hexString;
}

@end
