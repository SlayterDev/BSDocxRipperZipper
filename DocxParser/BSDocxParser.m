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

-(void) loadString {
	// get body element
    GDataXMLElement *body = [[self.xmlDoc.rootElement elementsForName:@"w:body"] objectAtIndex:0];
    
	NSArray *paragraphs = [body elementsForName:@"w:p"]; // find all parapgraphs
    for (GDataXMLElement *paragraph in paragraphs) {
        NSLog(@"Found a paragraph!");
        
        NSArray *runs = [paragraph elementsForName:@"w:r"]; // find all runs in each paragraph
        for (GDataXMLElement *run in runs) {
            NSLog(@"Found a run!");
			// get the text if it exists
            GDataXMLElement *runTextElem = [[run elementsForName:@"w:t"] objectAtIndex:0];
            if (runTextElem) {
				// Create attr string from text in run
                runString = [[NSMutableAttributedString alloc] initWithString:runTextElem.stringValue attributes:@{NSFontAttributeName: normalFont}];
                
				// Get run preferences
                GDataXMLElement *rPr = [[run elementsForName:@"w:rPr"] objectAtIndex:0];
				
				// Pull out bold attribute tag
                GDataXMLElement *boldTrait = [[rPr elementsForName:@"w:b"] objectAtIndex:0];
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
                
                [finalString appendAttributedString:runString];
                NSLog(@"Final: %@", finalString.string);
            }
        }
        
        [[finalString mutableString] appendString:@"\n"]; // add a new line after each paragraph
    }
    
    NSLog(@"%@", finalString.string);
}

-(void) checkBoldWithElement:(GDataXMLElement *)element {
	if (element.XMLString && [element.XMLString hasSuffix:XML_TAG_TRUE]) {
		NSLog(@"SET BOLD YES");
		// Give the string bold attributes
		[runString addAttributes:@{NSFontAttributeName: boldFont} range:[runString fullRange]];
		NSLog(@"Run: %@", runString.string);
	}
}

-(void) checkItalicWithElement:(GDataXMLElement *)element {
	if (element.XMLString && [element.XMLString hasSuffix:XML_TAG_TRUE]) {
		NSLog(@"SET ITALIC YES");
		// Give string italic attributes
		[runString addAttributes:@{NSFontAttributeName: italicFont} range:[runString fullRange]];
	}
}

-(void) checkUnderlineWithElement:(GDataXMLElement *)element {
	if (element.XMLString && [element.XMLString hasSuffix:XML_UNDERLINE_TRUE_SINGLE]) {
		NSLog(@"SET UNDERLINE YES");
		// Give string italic attributes
		[runString addAttributes:@{NSUnderlineStyleAttributeName: [NSNumber numberWithInt:NSUnderlineStyleSingle]} range:[runString fullRange]];
	}
}

-(void) checkColorWithElement:(GDataXMLElement *)element {
	if (element.XMLString) {
		NSLog(@"Color attr: %@", element.attributes);
		GDataXMLNode *colorVal = [element attributeForName:@"w:val"];
		NSLog(@"Color val: %@", colorVal.stringValue);
		
		UIColor *fontColor = [UIColor colorWithHexString:colorVal.stringValue];
		[runString addAttributes:@{NSForegroundColorAttributeName: fontColor} range:[runString fullRange]];
	}
}

-(NSAttributedString *) getFinalString {
    return finalString;
}

@end
