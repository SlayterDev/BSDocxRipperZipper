//
//  BSFileHelper.m
//  TextEditor
//
//  Created by Brad Slayter on 5/22/12.
//  Copyright (c) 2012 __MyCompanyName__. All rights reserved.
//

#import "BSFileHelper.h"

@implementation BSFileHelper

static BSFileHelper *_sharedHelper;

+(BSFileHelper *) sharedHelper {
	if (_sharedHelper != nil)
		return _sharedHelper;
	
	_sharedHelper = [[BSFileHelper alloc] init];
	return _sharedHelper;
}

-(NSString *) getDocumentsDirectory {
	NSArray *paths = NSSearchPathForDirectoriesInDomains
	(NSDocumentDirectory, NSUserDomainMask, YES);
	NSString *documentsDirectory = [paths objectAtIndex:0];
	return documentsDirectory;
}

-(NSURL *) documentsDirectoryURL {
    return [NSURL fileURLWithPath:[self getDocumentsDirectory]];
}

-(void) createDirectoryWithName:(NSString *)name {
	NSString *documentsDirectory = [self getDocumentsDirectory];
	
	if (![[NSFileManager defaultManager] fileExistsAtPath:[NSString stringWithFormat:@"%@/%@", documentsDirectory, name]]) {
		
		[[NSFileManager defaultManager] createDirectoryAtPath:[NSString stringWithFormat:@"%@/%@", documentsDirectory, name] withIntermediateDirectories:YES attributes:nil error:nil];
		
	}
}

-(NSString *) getDirectoryPathWithName:(NSString *)name {
	NSString *documentsDirectory = [self getDocumentsDirectory];
	
	return [NSString stringWithFormat:@"%@/%@", documentsDirectory, name];
}

-(NSString *) getDataPlistPath {
	NSString *documentsDirectory = [self getDocumentsDirectory];
	
	return [NSString stringWithFormat:@"%@/data.plist", documentsDirectory];
}

-(void) writeFileWithName:(NSString *)name andText:(NSString *)text userFile:(BOOL)userFile {
	NSString *documentsDirectory = [self getDocumentsDirectory];
	NSString *path;
	
	if (userFile) {
		path = [NSString stringWithFormat:@"%@/files/%@.txt", documentsDirectory, name];
	} else {
		path = [NSString stringWithFormat:@"%@/%@.txt", documentsDirectory, name];
	}
	
	[text writeToFile:path atomically:YES encoding:NSStringEncodingConversionAllowLossy error:nil];
}

-(NSString *) loadFileForName:(NSString *)name {
	// **** BE SURE TO ASSIGN kCurrentFilename WHEN CALLED ****
	
	NSString *documentsDirectory = [self getDocumentsDirectory];
	NSString *path = [NSString stringWithFormat:@"%@/files/%@.txt", documentsDirectory, name];
	
	NSString *fileContents = [NSString stringWithContentsOfFile:path encoding:NSStringEncodingConversionAllowLossy error:nil];
	
	return fileContents;
}

-(void) deleteFileForName:(NSString *)name {
	NSString *documentsDirectory = [self getDocumentsDirectory];
	NSString *path = [NSString stringWithFormat:@"%@/files/%@", documentsDirectory, name];
	
	[[NSFileManager defaultManager] removeItemAtPath:path error:nil];
}

-(void) renameFile:(NSString *)fileName toNewName:(NSString *)newName {
	NSString *documentsDirectory = [self getDocumentsDirectory];
	NSString *path = [NSString stringWithFormat:@"%@/files/%@.txt", documentsDirectory, fileName];
	NSString *newPath = [NSString stringWithFormat:@"%@/files/%@.txt", documentsDirectory, newName];
	
	[[NSFileManager defaultManager] moveItemAtPath:path toPath:newPath error:nil];
	fileName = [fileName stringByAppendingString:@".txt"];
	[self deleteFileForName:fileName];
}

-(BOOL) fileExistsWithName:(NSString *)name {
	// ****	REQUIRES FILE EXTENSION AND DIRECTORY ****
	// unless you want to figure out how to automate that o_0
	
	NSString *documentsDirectory = [self getDocumentsDirectory];
	
	if ([[NSFileManager defaultManager] fileExistsAtPath:[NSString stringWithFormat:@"%@/%@", documentsDirectory, name]])
		return YES;
	else
		return NO;
}

@end
