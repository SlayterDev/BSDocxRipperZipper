//
//  BSFileHelper.h
//  TextEditor
//
//  Created by Brad Slayter on 5/22/12.
//  Copyright (c) 2012 __MyCompanyName__. All rights reserved.
//

#import <Foundation/Foundation.h>

@interface BSFileHelper : NSObject

+(BSFileHelper *) sharedHelper;
-(NSString *) getDocumentsDirectory;
-(NSURL *) documentsDirectoryURL;
-(void) createDirectoryWithName:(NSString *)name;
-(NSString *) getDirectoryPathWithName:(NSString *)name;
-(NSString *) getDataPlistPath;
-(void) writeFileWithName:(NSString *)name andText:(NSString *)text userFile:(BOOL)userFile;
-(BOOL) fileExistsWithName:(NSString *)name;
-(NSString *) loadFileForName:(NSString *)name;
-(void) deleteFileForName:(NSString *)name;
-(void) renameFile:(NSString *)fileName toNewName:(NSString *)newName;

@end
