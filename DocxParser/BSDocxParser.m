//
//  BSDocxParser.m
//  DocxParser
//
//  Created by Brad Slayter on 11/5/13.
//  Copyright (c) 2013 Brad Slayter. All rights reserved.
//

#import "BSDocxParser.h"

#define XML_TAG_TRUE			  @"w:val=\"1\"/>"
#define XML_UNDERLINE_TRUE_SINGLE @"w:val=\"single\"/>"

@interface NSMutableAttributedString (RangeExtension)
- (NSRange)fullRange;
@end

@implementation NSMutableAttributedString (RangeExtension)
- (NSRange)fullRange {
	return (NSRange){0, self.string.length};
}
@end

@implementation BSDocxParser

-(id) initWithFileURL:(NSURL *)fileURL {
    if ((self = [super init])) {
        self.fileURL = fileURL;
        
        [self setupFontDescriptor];
    }
    return self;
}

-(void) setupFontDescriptor {
    fontDescriptor = [UIFontDescriptor preferredFontDescriptorWithTextStyle:UIFontTextStyleBody];
    boldDescriptor = [fontDescriptor fontDescriptorWithSymbolicTraits:UIFontDescriptorTraitBold];
	italicDescriptor = [fontDescriptor fontDescriptorWithSymbolicTraits:UIFontDescriptorTraitItalic];
	
    boldFont = [UIFont fontWithDescriptor:boldDescriptor size:0.0];
	italicFont = [UIFont fontWithDescriptor:italicDescriptor size:0.0];
    normalFont = [UIFont preferredFontForTextStyle:UIFontTextStyleBody];
}

-(void) loadDocument {
    NSData *xmlData = [[NSMutableData alloc] initWithContentsOfURL:self.fileURL];
    NSError *error;
    _xmlDoc = [[GDataXMLDocument alloc] initWithData:xmlData options:0 error:&error];
    
    if (!self.xmlDoc) {
        NSLog(@"Couldn't open");
        return;
    }
    
    finalString = [[NSMutableAttributedString alloc] initWithString:@""];
    [self loadString];
}

-(void) resetHelperVars {
	boldFont = [UIFont fontWithDescriptor:boldDescriptor size:0.0];
	italicFont = [UIFont fontWithDescriptor:italicDescriptor size:0.0];
    normalFont = [UIFont preferredFontForTextStyle:UIFontTextStyleBody];
	
	currentFont = nil;
	currentFontName = nil;
	fontSize = 14.0f;
	runIsBold = NO;
	runIsItalic = NO;
}

-(void) loadString {
	// get body element
    GDataXMLElement *body = [[self.xmlDoc.rootElement elementsForName:@"w:body"] firstObject];
    
	NSArray *paragraphs = [body elementsForName:@"w:p"]; // find all parapgraphs
    for (GDataXMLElement *paragraph in paragraphs) {
        NSLog(@"Found a paragraph!");
		
		// check alignment
		alignmentAttr = nil;
		GDataXMLElement *pPr = [[paragraph elementsForName:@"w:pPr"] firstObject];
		GDataXMLElement *justificationElem = [[pPr elementsForName:@"w:jc"] firstObject];
		[self checkAlignmentWithElement:justificationElem];
        
        NSArray *runs = [paragraph elementsForName:@"w:r"]; // find all runs in each paragraph
        for (GDataXMLElement *run in runs) {
            NSLog(@"Found a run!");
			// get the text if it exists
            GDataXMLElement *runTextElem = [[run elementsForName:@"w:t"] firstObject];
            if (runTextElem) {
				// Create attr string from text in run
                runString = [[NSMutableAttributedString alloc] initWithString:runTextElem.stringValue attributes:@{NSFontAttributeName: normalFont}];
                
				// Get run preferences
                GDataXMLElement *rPr = [[run elementsForName:@"w:rPr"] firstObject];
				
				[self resetHelperVars];
				
				// Get font name
				GDataXMLElement *fontTrait = [[rPr elementsForName:@"w:rFonts"] firstObject];
				[self checkFontNameWithElement:fontTrait];
				
				// Get font size
				GDataXMLElement *sizeTrait = [[rPr elementsForName:@"w:sz"] firstObject];
				[self setFontSizeWithElement:sizeTrait];
				
				// Pull out bold attribute tag
                GDataXMLElement *boldTrait = [[rPr elementsForName:@"w:b"] firstObject];
                [self checkBoldWithElement:boldTrait];
				
				// Pull out italic tag
				GDataXMLElement *italicTrait = [[rPr elementsForName:@"w:i"] firstObject];
				[self checkItalicWithElement:italicTrait];
				
				// Pull out underline tag
				GDataXMLElement *underlineTrait = [[rPr elementsForName:@"w:u"] firstObject];
				[self checkUnderlineWithElement:underlineTrait];
				
				// Get color and set it
				GDataXMLElement *colorTrait = [[rPr elementsForName:@"w:color"] firstObject];
				[self checkColorWithElement:colorTrait];
				
				// apply custom font
				if (currentFont) {
					[runString addAttributes:@{NSFontAttributeName: currentFont} range:[runString fullRange]];
				}
				
				// apply alignment if needed
				if (alignmentAttr) {
					[runString addAttributes:alignmentAttr range:[runString fullRange]];
				}
                
                [finalString appendAttributedString:runString];
                NSLog(@"Run: %@", runString.string);
            }
        }
        
        [[finalString mutableString] appendString:@"\n"]; // add a new line after each paragraph
    }
    
    NSLog(@"%@", finalString.string);
}

-(void) checkFontNameWithElement:(GDataXMLElement *)element {
	if (element.XMLString) {
		GDataXMLNode *fontNode = [element attributeForName:@"w:ascii"];
		currentFont = [UIFont fontWithName:fontNode.stringValue size:fontSize];
		currentFontName = fontNode.stringValue;
		NSLog(@"Custom Font: %@", fontNode.stringValue);
	}
}

-(void) setFontSizeWithElement:(GDataXMLElement *)element {
	if (element.XMLString) {
		GDataXMLNode *sizeNode = [element attributeForName:@"w:val"];
		float size = (float)sizeNode.stringValue.intValue;
		size /= 2.0f;
		if (currentFont) {
			fontSize = size;
			currentFont = [UIFont fontWithName:currentFontName size:fontSize];
		} else {
			normalFont = [normalFont fontWithSize:size];
			boldFont = [boldFont fontWithSize:size];
			italicFont = [italicFont fontWithSize:size];
		}
	}
}

-(void) checkBoldWithElement:(GDataXMLElement *)element {
	if (element.XMLString && [element.XMLString hasSuffix:XML_TAG_TRUE]) {
		NSLog(@"SET BOLD YES");
		runIsBold = YES;
		// Give the string bold attributes
		if (!currentFont) {
			[runString addAttributes:@{NSFontAttributeName: boldFont} range:[runString fullRange]];
		} else {
			// custom font
			NSLog(@"%@", [self getNewFontNameForFont:currentFont.fontName bold:runIsBold italic:runIsItalic]);
			currentFont = [UIFont fontWithName:[self getNewFontNameForFont:currentFontName bold:runIsBold italic:runIsItalic] size:fontSize];
		}
		NSLog(@"Run: %@", runString.string);
		
	}
}

-(void) checkItalicWithElement:(GDataXMLElement *)element {
	if (element.XMLString && [element.XMLString hasSuffix:XML_TAG_TRUE]) {
		NSLog(@"SET ITALIC YES");
		runIsItalic = YES;
		// Give string italic attributes
		if (!currentFont) {
			[runString addAttributes:@{NSFontAttributeName: italicFont} range:[runString fullRange]];
		} else {
			// custom font
			currentFont = [UIFont fontWithName:[self getNewFontNameForFont:currentFontName bold:runIsBold italic:runIsItalic] size:fontSize];
		}
	}
}

-(void) checkUnderlineWithElement:(GDataXMLElement *)element {
	if (element.XMLString && [element.XMLString hasSuffix:XML_UNDERLINE_TRUE_SINGLE]) {
		NSLog(@"SET UNDERLINE YES");
		// Give string italic attributes
		[runString addAttributes:@{NSUnderlineStyleAttributeName: [NSNumber numberWithInt:NSUnderlineStyleSingle]} range:[runString fullRange]];
	}
}

-(void) checkAlignmentWithElement:(GDataXMLElement *)element {
	if (element.XMLString) {
		GDataXMLNode *alignNode = [element attributeForName:@"w:val"];
		NSMutableParagraphStyle *paragraphStyle = [[NSMutableParagraphStyle alloc] init];
		
		if ([alignNode.stringValue isEqualToString:@"center"]) {
			paragraphStyle.alignment = NSTextAlignmentCenter;
		} else if ([alignNode.stringValue isEqualToString:@"right"]) {
			paragraphStyle.alignment = NSTextAlignmentRight;
		} else if ([alignNode.stringValue isEqualToString:@"left"]) {
			paragraphStyle.alignment = NSTextAlignmentLeft;
		} else if ([alignNode.stringValue isEqualToString:@"justified"]) {
			paragraphStyle.alignment = NSTextAlignmentJustified;
		}
		
		alignmentAttr = @{NSParagraphStyleAttributeName: paragraphStyle};
	}
}

-(void) checkColorWithElement:(GDataXMLElement *)element {
	if (element.XMLString) {
		GDataXMLNode *colorVal = [element attributeForName:@"w:val"];
		NSLog(@"Color val: %@", colorVal.stringValue);
		
		UIColor *fontColor = [UIColor colorWithHexString:colorVal.stringValue];
		[runString addAttributes:@{NSForegroundColorAttributeName: fontColor} range:[runString fullRange]];
	}
}

-(NSString *) getNewFontNameForFont:(NSString *)fontName bold:(BOOL)bold italic:(BOOL)italic {
	NSArray *fonts = [UIFont fontNamesForFamilyName:fontName];
	NSLog(@"New font for bold or italic\n%@", fontName);
	
	for (NSString *font in fonts) {
		if ([font rangeOfString:@"bold" options:NSCaseInsensitiveSearch].location != NSNotFound && bold && [font rangeOfString:@"italic" options:NSCaseInsensitiveSearch].location == NSNotFound) {
			NSLog(@"Custom font bold");
			return font;
		} else if ([font rangeOfString:@"italic" options:NSCaseInsensitiveSearch].location != NSNotFound && italic && [font rangeOfString:@"bold" options:NSCaseInsensitiveSearch].location == NSNotFound) {
			NSLog(@"Custom font italic");
			return font;
		} else if ([font rangeOfString:@"italic" options:NSCaseInsensitiveSearch].location != NSNotFound && [font rangeOfString:@"bold" options:NSCaseInsensitiveSearch].location != NSNotFound && bold && italic) {
			NSLog(@"Custom font bold and italic");
			return font;
		}
	}
	
	return NULL;
}

-(NSAttributedString *) getFinalString {
    return finalString;
}

@end
