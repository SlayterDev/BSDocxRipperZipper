//
//  BSDocxParser.m
//  DocxParser
//
//  Created by Brad Slayter on 11/5/13.
//  Copyright (c) 2013 Brad Slayter. All rights reserved.
//

#import "BSDocxParser.h"

@implementation BSDocxParser

-(id) initWithFileURL:(NSURL *)fileURL {
    if ((self = [super init])) {
        self.fileURL = fileURL;
    }
    return self;
}

-(void) loadDocument {
    NSData *xmlData = [[NSMutableData alloc] initWithContentsOfURL:self.fileURL];
    NSError *error;
    _xmlDoc = [[GDataXMLDocument alloc] initWithData:xmlData options:0 error:&error];
    
    if (!self.xmlDoc) {
        NSLog(@"Couldn't open");
        return;
    }
    
    //NSLog(@"%@", self.xmlDoc.rootElement);
    finalString = @"";
    [self loadString];
}

-(void) loadString {
    GDataXMLElement *body = [[self.xmlDoc.rootElement elementsForName:@"w:body"] objectAtIndex:0]; // get body element
    NSArray *paragraphs = [body elementsForName:@"w:p"]; // find all parapgraphs
    for (GDataXMLElement *paragraph in paragraphs) {
        NSLog(@"Found a paragraph!");
        
        NSArray *runs = [paragraph elementsForName:@"w:r"]; // find all runs in each paragraph
        for (GDataXMLElement *run in runs) {
            NSLog(@"Found a run!");
            GDataXMLElement *runTextElem = [[run elementsForName:@"w:t"] objectAtIndex:0];   // get the text of it exists
            if (runTextElem)
                finalString = [finalString stringByAppendingString:runTextElem.stringValue]; // add text from run to final string
        }
        
        finalString = [finalString stringByAppendingString:@"\n"]; // add a new line after each paragraph
    }
    
    NSLog(@"%@", finalString);
}

-(NSString *) getFinalString {
    return finalString;
}

@end
