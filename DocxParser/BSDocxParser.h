//
//  BSDocxParser.h
//  DocxParser
//
//  Created by Brad Slayter on 11/5/13.
//  Copyright (c) 2013 Brad Slayter. All rights reserved.
//

#import <Foundation/Foundation.h>
#import "GDataXMLNode.h"

@interface BSDocxParser : NSObject {
    NSString *finalString;
}

@property (nonatomic, strong) NSURL *fileURL;
@property (nonatomic, strong) GDataXMLDocument *xmlDoc;

-(id) initWithFileURL:(NSURL *)fileURL;
-(void) loadDocument;
-(NSString *) getFinalString;

@end
