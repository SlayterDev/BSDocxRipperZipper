//
//  BSDocxRipperZipper.h
//  DocxParser
//
//  Created by Bradley Slayter on 11/8/13.
//  Copyright (c) 2013 Brad Slayter. All rights reserved.
//

#import <Foundation/Foundation.h>
#import "SSZipArchive.h"
#import "BSFileHelper.h"
#import "BSDocxParser.h"
#import "BSDocxWriter.h"

@interface BSDocxRipperZipper : NSObject

@property (nonatomic, strong) NSURL *docxURL;
@property (nonatomic, strong) NSURL *xmlURL;

-(BSDocxParser *) openDocxAtURL:(NSURL *)fileURL;
-(void) writeStringToDocx:(NSAttributedString *)string;

@end
