//
//  ViewController.m
//  DocxParser
//
//  Created by Brad Slayter on 11/5/13.
//  Copyright (c) 2013 Brad Slayter. All rights reserved.
//

#import "ViewController.h"

@interface ViewController ()

@end

@implementation ViewController

- (void)viewDidLoad
{
    [super viewDidLoad];
	// Do any additional setup after loading the view, typically from a nib.
    
    BSDocxParser *parser = [[BSDocxParser alloc] initWithFileURL:[self getDocURL]];
    [parser loadDocument];
    
    self.textView.attributedText = [parser getFinalString];
    self.textView.dataDetectorTypes = UIDataDetectorTypeLink;
}

-(NSURL *) getDocURL {
    NSString *docxPath = [[NSBundle mainBundle] pathForResource:@"OP" ofType:@"docx"];
    NSString *zipPath = [[[BSFileHelper sharedHelper] getDocumentsDirectory] stringByAppendingPathComponent:@"OP.zip"];
    NSString *unZipPath = [[[BSFileHelper sharedHelper] getDocumentsDirectory] stringByAppendingPathComponent:@"OP"];
    NSError *error;
    [[NSFileManager defaultManager] moveItemAtPath:docxPath toPath:zipPath error:&error];
    if (error) {
        NSLog(@"Error: %@", error.localizedDescription);
    } else {
        NSLog(@"success");
    }
    
    [SSZipArchive unzipFileAtPath:zipPath toDestination:unZipPath];
    
    NSURL *docURL = [[BSFileHelper sharedHelper] documentsDirectoryURL];
    docURL = [docURL URLByAppendingPathComponent:@"OP/word/document.xml"];
    
    //return [[NSBundle mainBundle] URLForResource:@"document7" withExtension:@"xml"];
    return docURL;
}

- (void)didReceiveMemoryWarning
{
    [super didReceiveMemoryWarning];
    // Dispose of any resources that can be recreated.
}

@end
