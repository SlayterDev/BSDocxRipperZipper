//
//  BSDocxRipperZipper.m
//  DocxParser
//
//  Created by Bradley Slayter on 11/8/13.
//  Copyright (c) 2013 Brad Slayter. All rights reserved.
//

#import "BSDocxRipperZipper.h"

@implementation BSDocxRipperZipper

-(BSDocxParser *) openDocxAtURL:(NSURL *)fileURL {
	BSDocxParser *parser = [[BSDocxParser alloc] initWithFileURL:[self getDocURL:fileURL]];
	self.docxURL = fileURL;
	
	return parser;
}

-(NSURL *) getDocURL:(NSURL *)fileURL {
	NSString *docxPath = [fileURL path];
    NSString *zipPath = [[[BSFileHelper sharedHelper] getDocumentsDirectory] stringByAppendingPathComponent:[[fileURL lastPathComponent] stringByDeletingPathExtension]];
	NSString *unZipPath = zipPath;
	zipPath = [zipPath stringByAppendingPathExtension:@"zip"];
    NSError *error;
    [[NSFileManager defaultManager] moveItemAtPath:docxPath toPath:zipPath error:&error];
    if (error) {
        NSLog(@"Error: %@", error.localizedDescription);
    } else {
        NSLog(@"success");
    }
    
    [SSZipArchive unzipFileAtPath:zipPath toDestination:unZipPath];
    
    NSURL *docURL = [[BSFileHelper sharedHelper] documentsDirectoryURL];
    docURL = [docURL URLByAppendingPathComponent:[NSString stringWithFormat:@"%@/word/document.xml", [fileURL.lastPathComponent stringByDeletingPathExtension]]];
    self.xmlURL = docURL;
	NSLog(@"Returning: %@", docURL);
	
    return docURL;
}

-(void) writeStringToDocx:(NSAttributedString *)string {
	BSDocxWriter *writer = [[BSDocxWriter alloc] initWithAttributedString:string];
	NSData *data = [writer buildDocument];
	[self writeXML:data];
}

-(void) writeXML:(NSData *)data {
	NSURL *filePath = self.xmlURL;
	[data writeToURL:filePath atomically:YES];
	
	NSString *dirPath = [[[BSFileHelper sharedHelper] getDocumentsDirectory] stringByAppendingPathComponent:[self.docxURL.lastPathComponent stringByDeletingPathExtension]];
	NSArray *paths = [[NSFileManager defaultManager] subpathsAtPath:dirPath];
	NSMutableArray *fullPaths = [[NSMutableArray alloc] init];
	for (int i = 0; i < paths.count; i++) {
		if ([[[paths objectAtIndex:i] lastPathComponent] hasPrefix:@".DS"])
			continue;
		
		[fullPaths addObject:[dirPath stringByAppendingPathComponent:[paths objectAtIndex:i]]];
	}
	
	NSString *zipPath = [[[BSFileHelper sharedHelper] getDocumentsDirectory] stringByAppendingPathComponent:[[self.docxURL lastPathComponent] stringByDeletingPathExtension]];
	zipPath = [zipPath stringByAppendingPathExtension:@"zip"];
	
	NSString *destPath = zipPath;
	NSString *docPath = [self.docxURL path];
	[SSZipArchive createZipFileAtPath:destPath withFilesAtPaths:fullPaths];
	NSError *error;
	//[[NSFileManager defaultManager] removeItemAtPath:docPath error:nil];
	[[NSFileManager defaultManager] moveItemAtPath:destPath toPath:docPath error:&error];
	if (error) {
		NSLog(@"Error creating docx: %@", error.localizedDescription);
	} else {
		NSLog(@"Success");
		NSLog(@"done %@", docPath);
		[self cleanUp];
	}
}

-(void) cleanUp {
	NSString *dirPath = [[[BSFileHelper sharedHelper] getDocumentsDirectory] stringByAppendingPathComponent:[self.docxURL.lastPathComponent stringByDeletingPathExtension]];
	[[NSFileManager defaultManager] removeItemAtPath:dirPath error:nil];
	self.docxURL = nil;
	self.xmlURL = nil;
}

@end
