//
//  BSDocxWriter.h
//  DocxParser
//
//  Created by Bradley Slayter on 11/7/13.
//  Copyright (c) 2013 Brad Slayter. All rights reserved.
//

#import <Foundation/Foundation.h>
#import "GDataXMLNode.h"
#include "UIColor+HexString.h"

@interface BSDocxWriter : NSObject {
	NSMutableArray *paragraphs;
}

@property (nonatomic, strong) NSAttributedString *string;

-(id) initWithAttributedString:(NSAttributedString *)string;

@end
