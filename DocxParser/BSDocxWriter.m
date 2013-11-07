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

-(void) splitString {
	NSMutableArray *tempParagraphs = [[NSMutableArray alloc] init];
	
	[self.string enumerateAttributesInRange:[self.string fullRange] options:0 usingBlock:^(NSDictionary *attrs, NSRange range, BOOL *stop) {
		
		NSString *subString = [[self.string string] substringWithRange:range];
		NSAttributedString *attrString = [[NSAttributedString alloc] initWithString:subString attributes:attrs];
		
		[tempParagraphs addObject:attrString];
									 
	}];
	
	paragraphs = [[NSMutableArray alloc] init];
	for (NSAttributedString *subString in tempParagraphs) {
		NSArray *subSubStrings = [[subString string] componentsSeparatedByString:@"\n"];
		
		for (NSString *s in subSubStrings) {
			NSAttributedString *attrString = [[NSAttributedString alloc] initWithString:s attributes:[subString attributesAtIndex:0 effectiveRange:NULL]];
			[paragraphs addObject:attrString];
			if (subSubStrings.count > 1) {
				NSAttributedString *newLineParagraph = [[NSAttributedString alloc] initWithString:@"\n" attributes:[subString attributesAtIndex:0 effectiveRange:NULL]];
				[paragraphs addObject:newLineParagraph];
			}
		}
	}
	
	NSLog(@"Printing");
	for (NSAttributedString *s in paragraphs) {
		NSLog(@"%@", s.string);
	}
	
	[self buildRootElement];
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

@end
