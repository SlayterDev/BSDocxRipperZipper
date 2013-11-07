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
}

@end
