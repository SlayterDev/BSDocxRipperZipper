//
//  BSDocxParser.h
//  DocxParser
//
//  Created by Brad Slayter on 11/5/13.
//  Copyright (c) 2013 Brad Slayter. All rights reserved.
//

#import <Foundation/Foundation.h>
#import "GDataXMLNode.h"
#include "UIColor+HexString.h"

@interface BSDocxParser : NSObject {
    NSMutableAttributedString *finalString;
	NSMutableAttributedString *runString;
    
    UIFontDescriptor *fontDescriptor;
    UIFontDescriptor *boldDescriptor;
	UIFontDescriptor *italicDescriptor;
	
    UIFont *boldFont;
    UIFont *normalFont;
	UIFont *italicFont;
	
	UIFont *currentFont;
	NSString *currentFontName;
	
	float fontSize;
	BOOL runIsBold;
	BOOL runIsItalic;
    BOOL runIsHyperlink;
	
	NSDictionary *alignmentAttr;
}

@property (nonatomic, strong) NSURL *fileURL;
@property (nonatomic, strong) GDataXMLDocument *xmlDoc;

-(id) initWithFileURL:(NSURL *)fileURL;
-(void) loadDocument;
-(NSAttributedString *) getFinalString;

@end
