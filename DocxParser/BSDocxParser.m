//
//  BSDocxParser.m
//  DocxParser
//
//  Created by Brad Slayter on 11/5/13.
//  Copyright (c) 2013 Brad Slayter. All rights reserved.
//

#import "BSDocxParser.h"

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
    boldFont = [UIFont fontWithDescriptor:boldDescriptor size:0.0];
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
                NSMutableAttributedString *runString = [[NSMutableAttributedString alloc] initWithString:runTextElem.stringValue attributes:@{NSFontAttributeName: normalFont}];
                
				// Pull out bold attribute tag
                GDataXMLElement *rPr = [[run elementsForName:@"w:rPr"] objectAtIndex:0];
                GDataXMLElement *boldTrait = [[rPr elementsForName:@"w:b"] objectAtIndex:0];
                if (boldTrait.XMLString) {
                    if ([boldTrait.XMLString hasSuffix:@"w:val=\"1\"/>"]) {
                        NSLog(@"SET BOLD YES");
						// Give the string bold attributes
                        [runString setAttributes:@{NSFontAttributeName: boldFont} range:[runString fullRange]];
                        NSLog(@"Run: %@", runString.string);
                    }
                }
                
                [finalString appendAttributedString:runString];
                NSLog(@"Final: %@", finalString.string);
            }
        }
        
        [[finalString mutableString] appendString:@"\n"]; // add a new line after each paragraph
    }
    
    NSLog(@"%@", finalString.string);
}

-(NSAttributedString *) getFinalString {
    return finalString;
}

@end
